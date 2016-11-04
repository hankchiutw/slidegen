// This file is required by the index.html file and will
// be executed in the renderer process for that window.
// All of the Node.js APIs are available in this process.

const {ipcRenderer} = require('electron')

const submitAnchor = document.getElementById('submitAnchor');
const message = document.getElementById('message');

submitAnchor.addEventListener('change', function(event){
    // validate the file
    const file = event.target.files[0];
    if(!file){
        log('未選擇檔案');
        return event.preventDefault();
    }

    if(file.type != 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'){
        log('檔案類型必須是docx');
        return event.preventDefault();
    }

    ipcRenderer.send('submitFile', file.path);
})

ipcRenderer.on('extracted', function(event, result){
    document.getElementById('rawText').innerHTML = result.rawText;
    document.getElementById('parsedText').innerHTML = result.parsedText;
})

ipcRenderer.on('doParsing', function(event, message){
    log(message)
})

ipcRenderer.on('generated', function(event, path){
    log(`pptx已產生：${path}`)
})

function log(str){
    message.innerHTML += `<p>${str}</p>`;
}
