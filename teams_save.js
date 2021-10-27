// find the collapsed replies and expand them
$('div.conversation-collapsed').find('div.expand-collapse:not(.chevron-expanded)').find('a.ts-collapsed-string').click();

// let the expansions load
await new Promise(r => setTimeout(r, 1000));
console.log('Expanding replies');

// find the "see more" button links and expand them
$('button.ts-see-more-fold').click();

// let the expansions load
await new Promise(r => setTimeout(r, 1000));
console.log('Expanding long messages');

function toMapable(fn) {
    //// Decorator function to pass a function to the $().map interface.
    function wrapper(ix, e) {
        return fn(e)
    }
    return wrapper
}

function getMessageData(message) {
    //// Gathers the header and body info from a single message in Teams
    let msg = jQuery(message);
    let msg_id = msg.attr('id');
    let header = msg.find('div.top-row-text-container');
    let name = header.find('div.ts-msg-name').text().trim();
    let date = header.find('span.ts-created.message-datetime').attr('title');
    let body_div = msg.find('div.message-body-content');
    let body = body_div.text().trim();
    let images = body_div.find('img.ts-image');
    let imgs_base64_src = images.map(toMapable(imgToImage64Src)).toArray();
    var data = {
        'msg_id': msg_id,
        'name': name,
        'date': date,
        'body': body,
        'images_base64_src': imgs_base64_src
    }
    return data
}

function imgToImage64Src(img) {
    let canvas = document.createElement('canvas');
    let context = canvas.getContext('2d');
    canvas.height = img.naturalHeight;
    canvas.width = img.naturalWidth;
    context.drawImage(img, 0, 0, img.naturalWidth, img.naturalHeight);
    let image64_src = canvas.toDataURL();
    return image64_src
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

    var reply_data = replies.map(toMapable(getMessageData));
    data['replies'] = reply_data.toArray();
    return data
}

function download(content, filename, contentType) {
    var a = document.createElement("a");
    var file = new Blob([content], {type: contentType});
    a.href = URL.createObjectURL(file);
    a.target = '_blank';
    a.download = filename;
    a.click();
}

function saveTeamsConversations() {
    //// Saves all messages in the current Team view in your browser

    // grab the list of message threads
    let conversation_list = $('div.ts-message-list-item');
    let data = conversation_list.map(toMapable(getConversationData));
    let json_data_str = JSON.stringify(data.toArray(), null, '  ');

    let now = new Date(Date.now());
    let timestamp = now.toISOString().split('.')[0].replaceAll(':', '.');
    let filename = 'Teams-json_' + timestamp + '.json';
    download(json_data_str, filename, 'application/json');
    return data
}

await new Promise(r => setTimeout(r, 2000));
console.log('Scraping conversations');
var data = saveTeamsConversations();