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

function imgToImage64Src(img) {
    let canvas = document.createElement('canvas');
    let context = canvas.getContext('2d');
    canvas.height = img.naturalHeight;
    canvas.width = img.naturalWidth;
    context.drawImage(img, 0, 0, img.naturalWidth, img.naturalHeight);
    let image64_src = canvas.toDataURL();
    return image64_src
}

function download(content, filename, contentType) {
    var a = document.createElement("a");
    var file = new Blob([content], {type: contentType});
    a.href = URL.createObjectURL(file);
    a.target = '_blank';
    a.download = filename;
    a.click();
}

function getMessageData(message_item) {
    //// Gathers the header and body info from a single message in Teams

    // row-level
    let message_thread = jQuery(message_item);
    let thread_id = message_thread.children('div.clearfix').attr('id');

    // bubble-level
    let message = message_thread.find('div.thread-body');
    let message_id = message.attr('id');

    // header
    let header = message.find('div.top-row-text-container'); 
    let name = header.find('div.ts-msg-name').text().trim();
    let date = header.find('span.ts-created.message-datetime').attr('title');

    // content
    let body_div = message.find('div.message-body-content');
    let body = body_div.text().trim();

    // image attachments
    let images = message.find('img.ts-image');
    let imgs_base64_src = images.map(toMapable(imgToImage64Src)).toArray();
    var data = {
        'thread_id': thread_id,
        'msg_id': message_id,
        'name': name,
        'date': date,
        'body': body,
        'images_base64_src': imgs_base64_src
    }
    return data
}

function saveChatData() {
    //// Save the messages from the "Chat" sidetab
    let message_list = $('div.ts-message-list-item');
    let data = message_list.map(toMapable(getMessageData));
    let json_data_str = JSON.stringify(data.toArray(), null, '  ');

    let now = new Date(Date.now());
    let timestamp = now.toISOString().split('.')[0].replaceAll(':', '.');
    let filename = 'Chat-json_' + timestamp + '.json';
    download(json_data_str, filename, 'application/json');
    return data
}

