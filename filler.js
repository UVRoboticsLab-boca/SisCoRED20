var http = require('http');
var formidable = require('formidable');
var mime = require('mime')
var PizZip = require('pizzip');
var Docxtemplater = require('docxtemplater');

var fs = require('fs');
var path = require('path');
var url = require('url');

http.createServer(function (req, res) {
 if (req.url.indexOf("/generate")!==-1){//== '/fileupload') {
    var form = new formidable.IncomingForm();
    form.parse(req, function (err, fields, files) {
      var content = fs
            .readFileSync(path.resolve(__dirname, './sources/formato.docx'), 'binary');
      
      var zip = new PizZip(content);
      var doc = new Docxtemplater();       
      console.log(fields.nomPE)
      //res.write('Haciendo Forma...');
      doc.loadZip(zip);
      if(fields.area=="prim"){
      	var area1="X";
      	var area2="";
      }
      else
      {
		var area1="";
      	var area2="X";
      }

      //set the templateVariables
      doc.setData({
          programa: fields.nomPE,
          campus: fields.nomCampus,
          dependencia: fields.nomDep,
          codigo: fields.codigo,
          experiencia: fields.nomEE,
          area_1: area1,
          area_2: area2
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

      var nombreArchivo = fields.nomEE.replace(/\s/g, '');
      if (nombreArchivo==""){
      	nombreArchivo="output.docx"
      }
      else{
      	nombreArchivo=nombreArchivo+".docx"
      }
      // buf is a nodejs buffer, you can either write it to a file or do anything else with it.
      fs.writeFileSync(path.resolve(__dirname, nombreArchivo), buf);
      
      //hacer download del archivo
      //var file = __dirname + '/output.docx';
	  //var filename = path.basename(file);
  	  //var mimetype = mime.lookup(file);

  
      console.log("Cargando...");
      try{
      	var archivo=nombreArchivo;
      	var ruta = path.resolve(__dirname, archivo)
  	  	var file = fs
            .readFileSync(ruta , 'binary');
  	  	res.setHeader('Content-disposition', 'attachment; filename=' + archivo);
      	res.setHeader('Content-type', "application/msword");
      	res.setHeader('Content-Length', file.length);
  		res.write(file, 'binary');
  		res.end();
  
      	
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



