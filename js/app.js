'use strict';

const {app, ipcMain} = require('electron')

const fs = require('fs');
const mammoth = require('mammoth');
const officegen = require('officegen');

/**
 * Extract text and generate the pptx when the input file submitted
 */
ipcMain.on('submitFile', function(event, filePath){
    console.log('start extracting:', filePath)

    mammoth.extractRawText({path: filePath})
        .then(function(result){
            console.log('extracted, raw result:', result);

            const pptx = initPptx();
            const words = result.value.split('\n\n\n\n\n\n')[0].split('\n\n\n\n');
            console.log('parsed words:', words);
            event.sender.send('extracted', { rawText: result.value, parsedText: words.toString() });
            event.sender.send('doParsing', `${words.length} sentences parsed.`);

            // add slides
            words.forEach(function(word){
                word.split('\n\n').forEach((subWord) => addWordSlide(pptx, subWord) );
                console.log('slide appended:', word);
                event.sender.send('doParsing', `產生投影片: ${word}`);
            });

            // write to the output file
            const outputPath = filePath.replace(filePath.substr(filePath.lastIndexOf('.')), '.pptx');
            const out = fs.createWriteStream(outputPath).on('error', console.log);
            pptx.generate(out);
            console.log('output file:', outputPath);
            event.sender.send('generated', outputPath);

        }).done();

});

/**
 * Initialize a pptx object
 */
function initPptx(){
    // init pptx
    const pptx = officegen('pptx');
    pptx.setWidescreen(true);

    pptx.on('finalize', function (written) {
      console.log('pptx finalized');
    });

    pptx.on('error', function (err) {
      console.log(err);
    });

    return pptx;
}

/**
 * Add one slide to the pptx
 */
function addWordSlide(pptx, word){

  // Let's create a new slide:
    let slide = pptx.makeNewSlide();
    slide.addImage('./images/bg.jpg', { y: 0, x: 0, cy: '100%', cx: '100%' });

    // Declare the default color to use on this slide:
    slide.color = '000000';

  // Basic way to add text string:
    slide.addText(word, {
        y: 20,
        x: 20,
        cx: '96.7%',
        font_face: '標楷體',
        font_size: 44
    });
}
