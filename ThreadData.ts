/**
 * Copyright 2020 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

import Utils from './utils';
import ThreadAction, {BooleanActionType, InboxActionType} from './ThreadAction';
import {SessionData} from './SessionData';

// Represents a message in a thread
const MAX_BODY_PROCESSING_LENGTH = 65535;

export class MessageData {

    private static parseAddresses(str: string): string[] {
        return str.toLowerCase().split(',').map(address => address.trim());
    }

    private static parseMailingList(message: GoogleAppsScript.Gmail.GmailMessage): string {
        const mailing_list = message.getHeader('Mailing-list').trim();
        if (!!mailing_list) {
            // E.x. "list xyz@gmail.com; contact xyz-admin@gmail.com"
            const parts = mailing_list.split(';');
            for (const part of parts) {
                const [type, address] = part.trim().split(/\s+/);
                Utils.assert(typeof address !== 'undefined', `Unexpected mailing list: ${mailing_list}`);
                if (type.trim() === 'list') {
                    return address;
                }
            }
        }
        const list_id = message.getHeader('List-ID').trim();
        if (!!list_id) {
            // E.x. "<mygroup.gmail.com>"
            let address = list_id;
            if (address.length > 0 && address.charAt(0) == '<') {
                address = address.substring(1);
            }
            if (address.length > 0 && address.charAt(-1)) {
                address = address.slice(0, -1);
            }
            return address;
        }
        return '';
    }

    private static parseSenders(message: GoogleAppsScript.Gmail.GmailMessage): string[] {
        const original_sender = message.getHeader('X-Original-Sender').trim();
        const sender = message.getHeader('Sender').trim();
        const senders: string[] = [];
        if (!!original_sender) {
            senders.push(original_sender);
        }
        if (!!sender) {
            senders.push(sender);
        }
        return senders;
    }

    public readonly from: string;
    public readonly to: string[];
    public readonly cc: string[];
    public readonly bcc: string[];
    public readonly list: string;
    public readonly reply_to: string[]; // TODO: support it in Rule
    public readonly sender: string[];
    public readonly receivers: string[];
    public readonly subject: string;
    public readonly body: string;
    public readonly headers: Map<string, string>;
    public readonly thread_labels: string[];
    public readonly thread_is_important: boolean;
    public readonly thread_is_in_inbox: boolean;
    public readonly thread_is_in_priority_inbox: boolean;
    public readonly thread_is_in_spam: boolean;
    public readonly thread_is_in_trash: boolean;
    public readonly thread_is_starred: boolean;
    public readonly thread_is_unread: boolean;
    public readonly thread_first_message_subject: string;

    constructor(session_data: SessionData, message: GoogleAppsScript.Gmail.GmailMessage) {
        this.from = message.getFrom();
        this.to = MessageData.parseAddresses(message.getTo());
        this.cc = MessageData.parseAddresses(message.getCc());
        this.bcc = MessageData.parseAddresses(message.getBcc());
        this.list = MessageData.parseMailingList(message);
        this.reply_to = MessageData.parseAddresses(message.getReplyTo());
        this.sender = ([] as string[]).concat(
            this.from, this.reply_to, ...MessageData.parseSenders(message));
        this.receivers = ([] as string[]).concat(this.to, this.cc, this.bcc, this.list);
        this.subject = message.getSubject();
        this.headers = new Map<string, string>();
        session_data.requested_headers.forEach(header => {
            this.headers.set(header, message.getHeader(header));
        });
        this.thread_labels = [];
        const thread = message.getThread();
        thread.getLabels().forEach(label => {
            this.thread_labels.push(label.getName());
        });
        this.thread_is_important = thread.isImportant();
        this.thread_is_in_inbox = thread.isInInbox();
        this.thread_is_in_priority_inbox = thread.isInPriorityInbox();
        this.thread_is_in_spam = thread.isInSpam();
        this.thread_is_in_trash = thread.isInTrash();
        this.thread_is_starred = thread.hasStarredMessages();
        this.thread_is_unread = thread.isUnread();
        this.thread_first_message_subject = thread.getFirstMessageSubject();
        // Potentially could be HTML, Plain, or RAW. But doesn't seem very useful other than Plain.
        let body = message.getPlainBody();
        // Truncate and log long messages.
        if (body.length > MAX_BODY_PROCESSING_LENGTH) {
            Logger.log(`Ignoring the end of long message with subject "${this.subject}"`);
            body = body.substring(0, MAX_BODY_PROCESSING_LENGTH);
        }
        this.body = body;
    }

    toString() {
        return this.subject;
    }
}

// Represents a thread
export class ThreadData {
    private readonly raw: GoogleAppsScript.Gmail.GmailThread;

    public readonly message_data_list: MessageData[];
    public readonly thread_action = new ThreadAction();

    constructor(session_data: SessionData, thread: GoogleAppsScript.Gmail.GmailThread) {
        this.raw = thread;

        const messages = thread.getMessages();
        // Get messages that is not too old, but at least one message
        let newMessages = messages.filter(
            message => message.getDate() > session_data.oldest_to_process);
        if (newMessages.length === 0) {
            newMessages = [messages[messages.length - 1]];
        }
        this.message_data_list = newMessages.map(message => new MessageData(session_data, message));

        // Log if any dropped.
        const numDropped = messages.length - newMessages.length;
        if (numDropped > 0) {
            const subject = this.message_data_list[0].subject;
            Logger.log(`Ignoring oldest ${numDropped} messages in thread "${subject}"`);
        }
    }

    getLatestMessage(): GoogleAppsScript.Gmail.GmailMessage {
        const messages = this.raw.getMessages();
        return messages[messages.length -1];
    }

    getFirstMessageSubject(): string {
        return this.raw.getFirstMessageSubject();
    }

    static applyAllActions(session_data: SessionData, all_thread_data: ThreadData[]) {
        const label_action_map: { [key: string]: GoogleAppsScript.Gmail.GmailThread[] } = {};
        const moving_action_map = new Map<InboxActionType, GoogleAppsScript.Gmail.GmailThread[]>([
            [InboxActionType.DEFAULT, []], [InboxActionType.INBOX, []], [InboxActionType.ARCHIVE, []], [InboxActionType.TRASH, []]
        ]);
        const important_action_map = new Map<BooleanActionType, GoogleAppsScript.Gmail.GmailThread[]>([
            [BooleanActionType.DEFAULT, []], [BooleanActionType.ENABLE, []], [BooleanActionType.DISABLE, []]
        ]);
        const read_action_map = new Map<BooleanActionType, GoogleAppsScript.Gmail.GmailThread[]>([
            [BooleanActionType.DEFAULT, []], [BooleanActionType.ENABLE, []], [BooleanActionType.DISABLE, []]
        ]);
        all_thread_data.forEach(thread_data => {
            const thread = thread_data.raw;
            const action = thread_data.thread_action;
            console.log(`apply action ${action} to thread '${thread.getFirstMessageSubject()}'`);

            // update label action map
            action.label_names.forEach(label_name => {
                if (!(label_name in label_action_map)) {
                    label_action_map[label_name] = [];
                }
                label_action_map[label_name].push(thread);
            });

            // other actions
            moving_action_map.get(action.move_to)!.push(thread);
            important_action_map.get(action.important)!.push(thread);
            read_action_map.get(action.read)!.push(thread);
        });

        Utils.withTimer("BatchApply", () => {
            // batch update labels
            for (const label_name in label_action_map) {
                const threads = label_action_map[label_name];
                session_data.getOrCreateLabel(label_name).addToThreads(threads);
                console.log(`add label ${label_name} to ${threads.length} threads`);
            }
            Logger.log(`Updated labels: ${Object.keys(label_action_map)}.`);

            moving_action_map.forEach((threads, action_type) => {
                switch (action_type) {
                    case InboxActionType.INBOX:
                        GmailApp.moveThreadsToInbox(threads);
                        break;
                    case InboxActionType.ARCHIVE:
                        GmailApp.moveThreadsToArchive(threads);
                        break;
                    case InboxActionType.TRASH:
                        GmailApp.moveThreadsToTrash(threads);
                        break;
                }
            });
            important_action_map.forEach((threads, action_type) => {
                switch (action_type) {
                    case BooleanActionType.ENABLE:
                        GmailApp.markThreadsImportant(threads);
                        break;
                    case BooleanActionType.DISABLE:
                        GmailApp.markThreadsUnimportant(threads);
                        break;
                }
            });
            read_action_map.forEach((threads, action_type) => {
                switch (action_type) {
                    case BooleanActionType.ENABLE:
                        GmailApp.markThreadsRead(threads);
                        break;
                    case BooleanActionType.DISABLE:
                        GmailApp.markThreadsUnread(threads);
                        break;
                }
            });
            Logger.log(`Updated threads status.`);

            const all_threads = all_thread_data.map(data => data.raw);
            if (session_data.config.processed_label.length > 0){
                session_data.getOrCreateLabel(session_data.config.processed_label).addToThreads(all_threads);
            }
            session_data.getOrCreateLabel(session_data.config.unprocessed_label).removeFromThreads(all_threads);
            Logger.log(`Mark as processed.`);
        });
    }
}
