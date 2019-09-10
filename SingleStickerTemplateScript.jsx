// Github:
// https://github.com/OHovey/WovenLabels-JustStick/blob/master/SingleStickerTemplateScript.jsx

var Doc = app.documents.item(0); 
var allGraphics = app.activeDocument.allGraphics
var pages = Doc.pages; 

// $.writeln('layers: ' + Doc.layers.item(0).allPageItems[0].label) 
// for (var item in Doc.layers.everyItem().allPageItems) {
//     $.writeln(item)
// }


// This array is used to compare a pages' image link name to, if there is a match, the page will go through
// extra alterations  
var noMotifImgNames = [
    'P100-B.eps',
    'P100-G.eps',
    'P100-PU.eps',
    'P100-R.eps',
    'P100-RB.eps',
    'P100-MB.eps',
    'P100-MG.eps'
]

// This Associative array is used to compare the colour component of a page link filename to the intended
// default colour of the page. The relevent value is then used to generate a color object from the ExtendScript apis' 
// list of preset values, which will then be used to set the color of the page text
var fillColorObjectMap = {
    'B': 'Black',
    'R': 'SM Red',
    'G': 'CMYK Green',
    'MG': 'Magenta',
    'MB': 'Cyan',
    'RB': 'Cyan',
    'PU': app.activeDocument.colors.item(3).name
}


// This function is called when the script determins that one of the TextFrames in the page 
// contains overset text. If so, it is nessarry to temporarily increase the size of the relevent text box 
// in order to access the 'pointSize' attribute of the paragraph text
//
//  Accepts:
//      i: Integer, index position of the page in the document
//
//  Returns: nothing
//
function reshapeAndResize(i) {
    try {
        // iterate through the text frames in the page
        for (var s = 0; s < textFrameCount; s++) {
            var page = pages.item(i + 1)

            // isolate the value of the current page bounds
            var origionalBounds = pages.item(i + 1).textFrames[s].geometricBounds
            // reset the geometric bounds of the text frame to those of the page bounds
            page.textFrames[s].geometricBounds = page.bounds 

            // Now that the target characters are availiable to target, reset the relavent attributes, pointSize and fillColor
            // page.textFrames.everyItem().paragraphs.everyItem().characters.everyItem().pointSize = 5
            page.textFrames.everyItem().paragraphs.everyItem().fillColor = myColor  

            // After the relavent attributes have been changed, reset the geometric bounds of the text frame to 
            // the origional bounds
            page.textFrames[s].geometricBounds = origionalBounds
        }
    } catch (err) {
        $.writeln('err: ' + err)
    }
}


var count = 0 // The count varaible allows the script to adjust the index position of successive image links in the document after 
              // they have been deleted from the script

// Begin iterating through the document pages
for (var i = 0; i <= pages.length; i++) {
    
    try {
        // start by identifying the colour identifier in the page image link
        var match = Doc.links.item(i - count).name.match(/-(\w)./)[0] 
        match = match.replace('-', '')
        match = match.replace('.', '')

        // use the color identifier to index the appropriate value from the 'fillColorObjectMap' above
        var myColorName = fillColorObjectMap[match]
        // select the page item (appears to be neccassary)
        pages.item(i + 1).select()
        // index the color object from the built in array of colors
        var myColor = app.activeDocument.colors.item(myColorName) 
        // reset the 'fillColor' property of the TextFrame paragraph text
        pages.item(i + 1).textFrames.everyItem().paragraphs.everyItem().fillColor = myColor

        // If any of the Text Frames contain overset text (which would mean the above did not work), reshape and resize the frame via the aptly named function above
        if (pages.item(i + 1).textFrames[0].overflows) reshapeAndResize(i)
        else if (pages.item(i + 1).textFrames[1].overflows) reshapeAndResize(i) 
        else if (pages.item(i + 1).textFrames[2].overflows) reshapeAndResize(i) 

    } catch (err) {
        switch(typeof err) {
            // If there is a runtime error, simply ignore and continue
            case 'RuntimeError':
                
        }
    }

    // start iterating through the number of items in the 'noMotifImageNames' array
    for (var s = 0; s < 7; s++) {
        try {
            // check if link name is equal to the indexed name
            if (Doc.links.item(i - count).name == noMotifImgNames[s]) {
                // index the correct page
                var page = pages.item(i + 1)
                try {
                    // if any of the page items are 'locked' and uneditable... change that
                    if (page.pageItems.everyItem().getElements()[0].locked) page.pageItems.everyItem().getElements()[0].locked = false
                    // Remove the 'no motif' filler image 
                    page.pageItems.everyItem().getElements()[0].remove()

                    // Given that descriminating between the target TextFrames is surprisingly tricky 
                    // all the text frames will be resized after the link image has been removed 
                    // here the 'pointSize' of the text frame can be altered as nessacery
                    var textFrameCount = page.textFrames.count() 
                    for (var t = 0; t < textFrameCount; t++) {
                        // page.textFrames.everyItem().paragraphs.everyItem().characters.everyItem().pointSize = 5 
                        page.textFrames[t].geometricBounds = page.bounds 
                    }
                    count++
                } catch (err) {
                    $.writeln('err: ' + err)
                }

                break;
            } 
        } catch (err) {
            switch(typeof err) {
                case 'RuntimeError': 
                    $.writeln('err: ' + err)
                    break;
            }
        }
    }
}
