var http = require('http');
var formidable = require('formidable');
var mime = require('mime')
var PizZip = require('pizzip');
var Docxtemplater = require('docxtemplater');

var fs = require('fs');
var path = require('path');
var url = require('url');

var conversorWord = function(cadena){
  var pre = '<w:p><w:r><w:t>';
  var post = '</w:t></w:r></w:p>';
  var lineBreak = '<w:br/>';
  var separados = pre + cadena.replace(/\n/g, lineBreak)+post;
  return separados;
};

http.createServer(function (req, res) {
 if (req.url.indexOf("/generate")!==-1){//== '/fileupload') {
    var form = new formidable.IncomingForm();
    form.parse(req, function (err, fields, files) {
      var content = fs
            .readFileSync(path.resolve(__dirname, './sources/formato.docx'), 'binary');
      
      var zip = new PizZip(content);
      var doc = new Docxtemplater(); 
      doc.setOptions({linebreaks:true});      
      console.log(fields.nomPE);
      //res.write('Haciendo Forma...');
      doc.loadZip(zip);
      
      console.log(fields);
      var campus ="Xalapa";
      var cp = parseInt(fields.nomCampus);
      switch (cp) {
        case 1:
          campus ="Veracruz";
          break;
        case 2:
              campus = "Orizaba - Córdoba";
          break;
        case 3:
              campus = "Coatzacoalcos - Minatitlán";
          break;
        case 4:
              campus = "Poza Rica - Tuxpan";
          break;
      }
    var depen ="";
    var dep = parseInt(fields.nomDep);

    switch(dep){
      case 0:
        depen="Arquitectura";
        break;
      case 1:
        depen="Ingeniería Civil";
        break;
      case 2:
        depen="Ingeniería Mecánica Eléctrica";
        break;
      case 3:
        depen="Ciencias Químicas";
        break;
      case 4:
        depen="Química Farmacéutica Bióloga";
        break;
      case 5:
        depen="Instrumentación Electrónica";
        break;
      case 6:
        depen="Matemáticas";
        break;
      case 7:
        depen="Física";
        break;
      case 8:
        depen="Ingeniería de la Construcción y el Hábitat";
        break;
      case 9:
        depen="Ingeniería Eléctrica y Electrónica";
        break;
      case 10:
        depen="Ingeniería Mecánica y Ciencias Navales";
        break;
      case 11:
        depen="Ingeniería";
        break;
      case 12:
        depen="Ingeniería Mecánica y Eléctrica";
        break;
      case 13:
        depen="Ingeniería en Electrónica y Comunicaciones";
        break;

    }

    var modalidad = "Curso";

    switch (fields.Modal) {
        case 1:
          modalidad ="Curso";
          break;
        case 2:
          modalidad = "Curso-Taller";
          break;
        case 3:
          modalidad = "Estancias académicas";
          break;
        case 4:
          modalidad = "Laboratorio";
          break;
        case 5:
          modalidad = "Práctica de campo";
          break;
        case 6:
          modalidad = "Práctica profesional";
          break;
        case 7:
          modalidad = "Seminario";
          break;
        case 8:
          modalidad = "Taller";
          break;
        case 9:
          modalidad = "Vinculación comunidad";
          break;
        case 10:
          modalidad = "Curso-Laboratorio";
          break;       

      }

      var estDeAprend = "";
      var estDeEns = "";

      if(fields.nomEstAp1){
        estDeAprend += "-" + fields.nomEstAp1 +"\n";
      }
      if(fields.nomEstAp2){
        estDeAprend += "-" + fields.nomEstAp2 +"\n";
      }
      if(fields.nomEstAp3){
        estDeAprend += "-" + fields.nomEstAp3 +"\n";
      }
      if(fields.nomEstAp4){
        estDeAprend += "-" + fields.nomEstAp4 +"\n";
      }
      if(fields.nomEstAp5){
        estDeAprend += "-" + fields.nomEstAp5 +"\n";
      }
      if(fields.nomEstAp5){
        estDeAprend += "-" + fields.nomEstAp6 +"\n";
      }
      if(fields.nomEstAp6){
        estDeAprend += "-" + fields.nomEstAp7 +"\n";
      }
      if(fields.nomEstAp7){
        estDeAprend += "-" + fields.nomEstAp8 +"\n";
      }
      if(fields.nomEstAp8){
        estDeAprend += "-" + fields.nomEstAp8 +"\n";
      }
      if(fields.nomEstAp9){
        estDeAprend += "-" + fields.nomEstAp9 +"\n";
      }
      if(fields.nomEstAp10){
        estDeAprend += "-" + fields.nomEstAp10 +"\n";
      }
      if(fields.nomEstAp12){
        estDeAprend += "-" + fields.nomEstAp12 +"\n";
      }
      if(fields.nomEstAp13){
        estDeAprend += "-" + fields.nomEstAp13 +"\n";
      }
      if(fields.nomEstAp14){
        estDeAprend += "-" + fields.nomEstAp14 +"\n";
      }
      if(fields.nomEstAp15){
        estDeAprend += "-" + fields.nomEstAp15 +"\n";
      }
      if(fields.nomEstAp16){
        estDeAprend += "-" + fields.nomEstAp16 +"\n";
      }
      if(fields.nomEstAp17){
        estDeAprend += "-" + fields.nomEstAp17 +"\n";
      }

      if(fields.nomEstApp1){
        estDeAprend += "-" + fields.nomEstApp1 +"\n";
      }
      if(fields.nomEstApp2){
        estDeAprend += "-" + fields.nomEstApp2 +"\n";
      }
      if(fields.nomEstApp3){
        estDeAprend += "-" + fields.nomEstApp3 +"\n";
      }
      if(fields.nomEstApp4){
        estDeAprend += "-" + fields.nomEstApp4 +"\n";
      }
      if(fields.nomEstApp5){
        estDeAprend += "-" + fields.nomEstApp5 +"\n";
      }
      if(fields.nomEstApp5){
        estDeAprend += "-" + fields.nomEstApp6 +"\n";
      }
      if(fields.nomEstApp6){
        estDeAprend += "-" + fields.nomEstApp7 +"\n";
      }
      if(fields.nomEstApp7){
        estDeAprend += "-" + fields.nomEstApp8 +"\n";
      }
      if(fields.nomEstApp8){
        estDeAprend += "-" + fields.nomEstApp8 +"\n";
      }
      if(fields.nomEstApp9){
        estDeAprend += "-" + fields.nomEstApp9 +"\n";
      }
      if(fields.nomEstApp10){
        estDeAprend += "-" + fields.nomEstApp10 +"\n"
      }
      if(fields.nomEstApp12){
        estDeAprend += "-" + fields.nomEstApp12 +"\n";
      }

      if(fields.nomEstApm1){
        estDeAprend += "-" + fields.nomEstApm1 +"\n";
      }
      if(fields.nomEstApm2){
        estDeAprend += "-" + fields.nomEstApm2 +"\n";
      }
      if(fields.nomEstApm3){
        estDeAprend += "-" + fields.nomEstApm3 +"\n";
      }
      if(fields.nomEstApm4){
        estDeAprend += "-" + fields.nomEstApm4 +"\n";
      }
      if(fields.nomEstApm5){
        estDeAprend += "-" + fields.nomEstApm5 +"\n";
      }
      if(fields.nomEstApm5){
        estDeAprend += "-" + fields.nomEstApm6 +"\n";
      }
      if(fields.nomEstApm6){
        estDeAprend += "-" + fields.nomEstApm7 +"\n";
      }
      if(fields.nomEstApm7){
        estDeAprend += "-" + fields.nomEstApm8 +"\n";
      }
      if(fields.nomEstApm8){
        estDeAprend += "-" + fields.nomEstApm8 +"\n";
      }
      if(fields.nomEstApm9){
        estDeAprend += "-" + fields.nomEstApm9 +"\n";
      }
      if(fields.nomEstApm10){
        estDeAprend += "-" + fields.nomEstApm10 +"\n";
      }
      if(fields.nomEstApm12){
        estDeAprend += "-" + fields.nomEstApm12 +"\n";
      }
      if(fields.nomEstApm13){
        estDeAprend += "-" + fields.nomEstApm13 +"\n";
      }
      if(fields.nomEstApm14){
        estDeAprend += "-" + fields.nomEstApm14 +"\n";
      }
      if(fields.nomEstApm15){
        estDeAprend += "-" + fields.nomEstApm15 +"\n";
      }
      if(fields.nomEstApm16){
        estDeAprend += "-" + fields.nomEstApm16 +"\n";
      }
      //**************************
      //estrategias de enseñanza
      //**************************
      if(fields.nomEstEnsT1){
        estDeEns += "-" + fields.nomEstEnsT1 +"\n";
      }
      if(fields.nomEstEnsT2){
        estDeEns += "-" + fields.nomEstEnsT2 +"\n";
      }
      if(fields.nomEstEnsT3){
        estDeEns += "-" + fields.nomEstEnsT3 +"\n";
      }
      if(fields.nomEstEnsT4){
        estDeEns += "-" + fields.nomEstEnsT4 +"\n";
      }
      if(fields.nomEstEnsT5){
        estDeEns += "-" + fields.nomEstEnsT5 +"\n";
      }
      if(fields.nomEstEnsT6){
        estDeEns += "-" + fields.nomEstEnsT6 +"\n";
      }

      if(fields.nomEstEnsP1){
        estDeEns += "-" + fields.nomEstEnsP1 +"\n";
      }
      if(fields.nomEstEnsP2){
        estDeEns += "-" + fields.nomEstEnsP2 +"\n";
      }
      if(fields.nomEstEnsP3){
        estDeEns += "-" + fields.nomEstEnsP3 +"\n";
      }
      
      if(fields.nomEstEnsM1){
        estDeEns += "-" + fields.nomEstEnsM1 +"\n";
      }
      if(fields.nomEstEnsM2){
        estDeEns += "-" + fields.nomEstEnsM2 +"\n";
      }
      if(fields.nomEstEnsM3){
        estDeEns += "-" + fields.nomEstEnsM3 +"\n";
      }
      if(fields.nomEstEnsM4){
        estDeEns += "-" + fields.nomEstEnsM4 +"\n";
      }
      if(fields.nomEstEnsM5){
        estDeEns += "-" + fields.nomEstEnsM5 +"\n";
      }
      if(fields.nomEstEnsM6){
        estDeEns += "-" + fields.nomEstEnsM6 +"\n";
      }

      //recursos didacticos
      var recursosDidac = "";
      if(fields.RecursoDidac1){
        recursosDidac += "-" + fields.RecursoDidac1 +"\n";
      }
      if(fields.RecursoDidac2){
        recursosDidac += "-" + fields.RecursoDidac2 +"\n";
      }
      if(fields.RecursoDidac3){
        recursosDidac += "-" + fields.RecursoDidac3 +"\n";
      }
      if(fields.RecursoDidac4){
        recursosDidac += "-" + fields.RecursoDidac4 +"\n";
      }
      if(fields.RecursoDidac5){
        recursosDidac += "-" + fields.RecursoDidac5 +"\n";
      }
      if(fields.RecursoDidac6){
        recursosDidac += "-" + fields.RecursoDidac6 +"\n";
      }
      if(fields.RecursoDidac7){
        recursosDidac += "-" + fields.RecursoDidac7 +"\n";
      }
      if(fields.RecursoDidac8){
        recursosDidac += "-" + fields.RecursoDidac8 +"\n";
      }
      if(fields.RecursoDidac9){
        recursosDidac += "-" + fields.RecursoDidac9 +"\n";
      }
      if(fields.RecursoDidac10){
        recursosDidac += "-" + fields.RecursoDidac10 +"\n";
      }
      if(fields.RecursoDidac11){
       recursosDidac += "-" + fields.RecursoDidac11 +"\n";
      }
      if(fields.RecursoDidac12!=""){
        recursosDidac += "-" + fields.RecursoDidac12 +"\n";
      }

      //Materiales 
      var MaterialesDidac="";

      if(fields.MatDidac1){
        MaterialesDidac += "-" + fields.MatDidac1 +"\n";
      }
      if(fields.MatDidac2){
        MaterialesDidac += "-" + fields.MatDidac2 +"\n";
      }
      if(fields.MatDidac3){
        MaterialesDidac += "-" + fields.MatDidac3 +"\n";
      }
      if(fields.MatDidac4){
        MaterialesDidac += "-" + fields.MatDidac4 +"\n";
      }
      if(fields.MatDidac5){
        MaterialesDidac += "-" + fields.MatDidac5 +"\n";
      }
      if(fields.MatDidac6){
        MaterialesDidac += "-" + fields.MatDidac6 +"\n";
      }
      if(fields.MatDidac7){
        MaterialesDidac += "-" + fields.MatDidac7 +"\n";
      }
      if(fields.MatDidac8){
        MaterialesDidac += "-" + fields.MatDidac8 +"\n";
      }
      if(fields.MatDidac9){
        MaterialesDidac += "-" + fields.MatDidac9 +"\n";
      }
      if(fields.MatDidac10){
        MaterialesDidac += "-" + fields.MatDidac10 +"\n";
      }
      if(fields.MatDidac11){
        MaterialesDidac += "-" + fields.MatDidac11 +"\n";
      }
      if(fields.MatDidac12){
        MaterialesDidac += "-" + fields.MatDidac12 +"\n";
      }
      if(fields.MatDidac13){
        MaterialesDidac += "-" + fields.MatDidac13 +"\n";
      }
      if(fields.MatDidac14){
        MaterialesDidac += "-" + fields.MatDidac14 +"\n";
      }
      if(fields.MatDidac15){
        MaterialesDidac += "-" + fields.MatDidac15 +"\n";
      }
      if(fields.MatDidac16){
        MaterialesDidac += "-" + fields.MatDidac16 +"\n";
      }
      if(fields.MatDidac17){
        MaterialesDidac += "-" + fields.MatDidac17 +"\n";
      }
      if(fields.MatDidac18){
        MaterialesDidac += "-" + fields.MatDidac18 +"\n";
      }
      if(fields.MatDidac19){
        MaterialesDidac += "-" + fields.MatDidac19 +"\n";
      }      
      if(fields.MatDidac20!=""){
        MaterialesDidac += "-" + fields.MatDidac20 +"\n";
      }

      var acred = "Para acreditar esta EE el estudiante deberá haber presentado con idoneidad y pertinencia cada evidencia de desempeño, es decir, que en cada una de ellas haya obtenido cuando menos el 60%.";
      var teoric = fields.nomSaberesT;//conversorWord(fields.nomSaberesT);
      var descripcion = fields.nomDesc + '\n' + fields.nomDesc2;
      var areasecu = "";
      if (fields.area2){
        areasecu=fields.area2;
      }
      //set the templateVariables
      doc.setData({
          programa: fields.nomPE,
          campus: campus,
          dependencia: depen,
          codigo: fields.codigo,
          experiencia: fields.nomEE,
          area_1: fields.area1,
          area_2: areasecu,
          creditos: fields.numCred,
          ht: fields.numHTeo,
          hp: fields.numHPra,
          th: fields.numTotHor,
          equivalencia: fields.nomEquiv,
          modal: modalidad,
          oportunidades: fields.oportu,
          prerequisitos: fields.nomPrereq,
          corequisitos: fields.nomCoreq,
          grupal: fields.nomInd,
          max: fields.nomMax,
          min: fields.nomMin,
          agrupacion: fields.nomAgrup,
          proyecto: fields.nomProy,
          felaboracion: fields.fecElab,
          fmodif: fields.fecModif,
          fapro: fields.fecAprob,
          academicos: fields.nomAcademic,
          perfil: fields.nomPerfil,
          espacio: fields.Espacio,
          relacion: fields.relacion,
          descrip: descripcion,
          justificacion: fields.nomJustificacion,
          unidad: fields.nomCompe,
          articulacion: fields.nomEjes,
          teoricos: teoric,
          heuristicos: fields.nomSaberesH,
          axiol: fields.nomSaberesA,
          estrategias1: estDeAprend,
          estrategias2: estDeEns,
          materiales: MaterialesDidac,
          recursos: recursosDidac,
          evidencia: fields.nomEvidencia,
          criterios: fields.nomCriterio,
          ambito: fields.nomAmbito,
          porcentaje: fields.nomPorcentaje,
          acreditacion: acred,
          fuentes: fields.nomFuentesBasicas,
          complementarias: fields.nomFuentesComp
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
      	nombreArchivo="output.docx";
      }
      else{
      	nombreArchivo=nombreArchivo+".docx";
      }

      //nombreArchivo = "./programas/" + nombreArchivo;
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



