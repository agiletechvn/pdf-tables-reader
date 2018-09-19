/* Copyright 2014 Mozilla Foundation
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */


var PAGE_NUMBER = 1;
var PAGE_SCALE = 1;

PDFJS.workerSrc = './build/pdf.worker.js';

function getText(marks, ex, ey) {
  var x = marks[0].x;
  var y = marks[0].y;
 
  var txt = '';
  var columns = 1;
  var lines = [];
  for (var i = 0; i < marks.length; i++) {
    var c = marks[i];
    var dx = c.x - x;
    var dy = c.y - y;
 
    if (Math.abs(dy) > ey) {
      // txt += "\"\n\"";
      lines.push({text:txt,x:c.x,y:c.y,columns:columns});
      txt = '';
      columns = 1;
      if (marks[i+1]) {
        // line feed - start from position of next line
        x = marks[i+1].x;
      }
    }
 
    if (Math.abs(dx) > ex && txt) {
      txt += "\",\"";
      columns++;
    }
 
    txt += c.c;
 
    x = c.x;
    y = c.y;
  }
 
  return lines;
}

function loadPDF(pdfAsArray) {
  PDFJS.getDocument(pdfAsArray).then(function (pdfDocument) {
    pdfDocument.getPage(PAGE_NUMBER).then(function (page) {
      var viewport = page.getViewport(PAGE_SCALE);
      console.log(viewport, page, pdfDocument)
      page.getTextContent().then(function (textContent) {
        // building SVG and adding that to the DOM
        // var svg = buildSVG(viewport, textContent);
        // document.getElementById('pageContainer').appendChild(svg);
        console.log(textContent)
        // should use filter for header
        var data = {
          'Kundenr': textContent.items[12].str,
          'Deres best. Dato': textContent.items[13].str,
          'Ordrebehandler': textContent.items[16].str,
          'E−mail': textContent.items[17].str
        }
        // window.items = textContent.items;
        // console.log(textContent.items)

        $.each(data, function(key){
          $(':contains("'+key+'")').prev().val(this).focus()
        })

        
        var chars = []

        // processing all items
        textContent.items.forEach(function (textItem) {
          // we have to take in account viewport transform, which includes scale,
          // rotation and Y-axis flip, and not forgetting to flip text.
          var tx = PDFJS.Util.transform(
            PDFJS.Util.transform(viewport.transform, textItem.transform),
            [1, 0, 0, -1, 0, 0]);    

          chars.push({
            x:tx[4],
            y:tx[5],
            c: textItem.str
          })
        });        

        chars.sort(
          function(a, b) {
            var dx = b.x - a.x;
            var dy = b.y - a.y;            
            if (Math.abs(dy) < 0.5) {
              return dx * -1;
            } else {
              return dy * -1;
            }
          }
        )
        // should use filter for column
        var lines = getText(chars, 1, 1);
        lines = lines.filter(function(line){
          var dx = Math.abs(line.x - 180);
          if(dx < 0.5 && line.columns == 5){            
            return true;
          }
          return false;
        }).map(function(line){            
          return line.text;
        });
        console.log(lines);
        $('#article').empty();
        lines.forEach(function(line){
          $('#article').append('<li>'+line.split(',')[0].replace(/\s*"$/,'')+'</li>')
        });
      });
    });
  });
}

document.addEventListener('DOMContentLoaded', function () {
  if (typeof PDFJS === 'undefined') {
    alert('Built version of PDF.js was not found.\n' +
          'Please run `gulp generic`.');
    return;
  }
  
  $('#done-button').on('click', function () {
    var file = $('#load-file')[0].files[0];
    var fileReader = new FileReader();
    fileReader.onloadend = function (e) {
      var arrayBuffer = e.target.result;
      loadPDF(arrayBuffer)
    };
    fileReader.readAsArrayBuffer(file);
  });

});