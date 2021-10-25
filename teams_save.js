

function getMessageData(message) {
    //// Gathers the header and body info from a single message in Teams
    let msg = jQuery(message);
    let msg_id = msg.attr('id');
    let header = msg.find('div.top-row-text-container');
    let name = header.find('div.ts-msg-name').text().trim();
    let date = header.find('span.ts-created.message-datetime').attr('title');
    let body = msg.find('div.message-body-content').text().trim();
    var data = {
        'msg_id': msg_id,
        'name': name,
        'date': date,
        'body': body
    }
    return data
}

function getConversationData(conversation_) {
    //// Gathers the header and body info from all messages in a conversation
    //// thread in Teams
    let conversation = jQuery(conversation_)
    let conversation_id = conversation.children('div.clearfix').attr('id');
    // primary message and replies in the conversation
    let primary_post = conversation.find('div.conversation-start').find('div.thread-body');
    let replies = conversation.find('div.conversation-reply').find('div.thread-body');

    // use the primary post as the return data template
    var data = getMessageData(primary_post);
    data['conversation_id'] = conversation_id;

    console.log(replies);
    var reply_data = replies.map(getMessageData);
    data['replies'] = reply_data.toArray();
    return data
}

function saveTeamsConversations() {
    //// Saves all messages in the current Team view in your browser

    // find the collapsed replies and expand them
    $('div.conversation-collapsed').find('div.expand-collapse:not(.chevron-expanded)').find('a.ts-collapsed-string').click();

    // grab the list of message threads
    let conversation_list = $('div.ts-message-list-item');
    var data = conversation_list.map(getConversationData);
    return data.toArray()
}