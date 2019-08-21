// open the template file if you must!
// app.open(File("/Users/Oliver/Documents/Motif Side Style/Single-Label4.indd"))

var Doc = app.documents.item(0); 
var allGraphics = app.activeDocument.allGraphics
var pages = Doc.pages; 

var noMotifImgNames = [
    'P100-B.eps',
    'P100-G.eps',
    'P100-PU.eps',
    'P100-R.eps',
    'P100-RB.eps',
    'P100-MB.eps',
    'P100-MG.eps'
]

var fillColorObjectMap = {
    'R': 'SM Red',
    'G': 'CMYK Green',
    'MG': 'Magenta',
    'MB': 'Cyan',
}


var count = 0
for (var i = 0; i <= pages.length; i++) {
    $.writeln('i: ' + i)
    
    try {
        var match = Doc.links.item(i - count).name.match(/-(\w)./)[0] 
        match = match.replace('-', '')
        match = match.replace('.', '')

        if (fillColorObjectMap[match] != undefined) {
            var myColorName = fillColorObjectMap[match]
            // $.writeln(i + ': ' + match + '===' + myColorName)
            pages.item(i + 1).select()
            var myColor = app.activeDocument.colors.item(myColorName) 
            pages.item(i + 1).textFrames.everyItem().paragraphs.everyItem().fillColor = myColor         
            
            // if (pages.item(i + 1).textFrames.item(1).overflows) {
            //     $.writeln(pages.item(i + 1).textFrames.everyItem().paragraphs.everyItem().characters.everyItem().pointSize)
            //     try {
            //         pages.item(i + 1).textFrames.everyItem().paragraphs.everyItem().characters.everyItem().pointSize = 4
            //     } catch (err) {
            //         $.writeln('err: ' + err)
            //     }
            // }
        }
    } catch (err) {
        switch(typeof err) {
            case 'RuntimeError':
                
        }
    }

    for (var s = 0; s < 7; s++) {
        try {
            if (Doc.links.item(i - count).name == noMotifImgNames[s]) {
                var page = pages.item(i + 1)
                try {
                    // $.writeln('pageIndex: ' + page.properties.name + '\n' + 'i: ' + i + '\n')
                    if (page.pageItems.everyItem().getElements()[0].locked) page.pageItems.everyItem().getElements()[0].locked = false
                    page.pageItems.everyItem().getElements()[0].remove()

                    var textFrameCount = page.textFrames.count() 
                    for (var s = 0; s < textFrameCount; s++) {
                        // CHANGE THIS TO TRANSFORM
                        page.textFrames[s].geometricBounds = page.bounds 
                    }
                    count++
                } catch (err) {
                    $.writeln('err: ' + err)
                }
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
