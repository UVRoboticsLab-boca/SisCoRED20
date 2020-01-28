var http = require('http');
var formidable = require('formidable');

var PizZip = require('pizzip');
var Docxtemplater = require('docxtemplater');

var fs = require('fs');
var path = require('path');
var url = require('url');
//Load the docx file as a binary


http.createServer(function (req, res) {
 if (req.url == '/fileupload') {
    var form = new formidable.IncomingForm();
    form.parse(req, function (err, fields, files) {
      var content = fs
            .readFileSync(path.resolve(__dirname, './sources/formato.docx'), 'binary');
      
      var zip = new PizZip(content);
      var doc = new Docxtemplater();       
      console.log(fields.nomPE)
      res.write('Haciendo Forma...');
      doc.loadZip(zip);

      //set the templateVariables
      doc.setData({
          programa: fields.nomPE,
          campus: fields.nomCampus,
          dependencia: fields.nomDep
      });

      try {
          // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
          doc.render()
      }
      catch (error) {
          var e = {
              message: error.message,
              name: error.name,
              stack: error.stack,
              properties: error.properties,
          }
          console.log(JSON.stringify({error: e}));
          // The error thrown here contains additional information when logged with JSON.stringify (it contains a property object).
          throw error;
      }

      var buf = doc.getZip()
                   .generate({type: 'nodebuffer'});

      // buf is a nodejs buffer, you can either write it to a file or do anything else with it.
      fs.writeFileSync(path.resolve(__dirname, 'output.docx'), buf);
      res.write('Forma Realizada');
      res.end();
    });
  } else {
      
     fs.readFile('form.html', function(err, data) {
        if (req.url.indexOf(".css")!==-1){
         console.log('Rendering a CSS');  
         res.writeHead(200, {'Content-Type': 'text/css'});
         var filename = path.join(__dirname, req.url);
         console.log("archivo:"+filename)
         res.write(filename);
        }
        else{
            res.writeHead(200, {'Content-Type': 'text/html'});   
            res.write(data);  
        }
       
       
       console.log(req.url);
       console.log(path.resolve(__dirname,"static/css/layouts/marketing.css"));
       res.end();
     });
  }
}).listen(8080);



