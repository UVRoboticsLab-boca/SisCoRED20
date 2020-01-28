var http = require('http');
var formidable = require('formidable');
var mime = require('mime')
var PizZip = require('pizzip');
var Docxtemplater = require('docxtemplater');

var fs = require('fs');
var path = require('path');
var url = require('url');
//Load the docx file as a binary
const mimetypes = {
    'html': 'text/html',
    'css': 'text/css',
    'js': 'text/javascript',
    'png': 'image/png',
    'jpeg': 'image/jpeg',
    'jpg': 'image/jpg',
    'ico': 'image/vnd'
};

http.createServer(function (req, res) {
 if (req.url.indexOf("/fileupload")!==-1){//== '/fileupload') {
    var form = new formidable.IncomingForm();
    form.parse(req, function (err, fields, files) {
      var content = fs
            .readFileSync(path.resolve(__dirname, './sources/formato.docx'), 'binary');
      
      var zip = new PizZip(content);
      var doc = new Docxtemplater();       
      console.log(fields.nomPE)
      //res.write('Haciendo Forma...');
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
      
      //hacer download del archivo
      var file = __dirname + '/output.docx';
	  //var filename = path.basename(file);
  	  //var mimetype = mime.lookup(file);

  
      console.log("Cargando...");
      try{
      	//var archivo="output.docx";
      	var archivo="planetas.jpg";
  	  	var ws = fs.createWriteStream('./Salida.jpg');
  	  	//res.setHeader('Content-disposition', 'attachment; filename=' + "Salida.jpg");
      	//res.setHeader('Content-type', "image/jpg");
  	  //res.write('Forma Realizada');
  	  	//var filestream = fs.createReadStream(path.resolve(__dirname, "output.docx"));
  	  	var filestream = fs.createReadStream(path.resolve(__dirname, archivo));
  	  	filestream.on("end", () => {
    		console.log("Stream ended");
  			res.setHeader('Content-disposition', 'attachment; filename=' + "Salida.jpg");
      		res.setHeader('Content-type', "image/jpg");
    		res.end();
		});
  	  	console.log(path.resolve(__dirname, archivo));
      	res.pipe(filestream);
      	filestream.pipe(ws);
      	//fs.createReadStream(path.join(__dirname, archivo)).pipe(ws);
    
      	//res.write(filestream);
      	
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
          //throw error;
          res.write('Error en  : '+(e.message));
          res.end();
      }
      });//fin parse

  } else {
      
     fs.readFile('form.html', function(err, data) {
        if (req.url.indexOf(".css")!==-1){
         console.log('Rendering a CSS');  
         res.writeHead(200, {'Content-Type': 'text/css'});
         var filename = path.join(__dirname, req.url);
         console.log("archivo: "+filename);
         var contents = fs.readFileSync(filename, 'utf8');
         res.write(contents);
        }
        else{
            res.writeHead(200, {'Content-Type': 'text/html'});   
            res.write(data);  
        }
       
       
       //console.log(req.url);
       //console.log(path.resolve(__dirname,"static/css/layouts/marketing.css"));
       res.end();
     });
  }
}).listen(8080);



