http = require('http');
const express = require('express');
const app = express();
const dirPath = require('path');

app.set('views', dirPath.join(__dirname,'views'));
app.engine('html', require('ejs').renderFile);
app.set('view engine', 'html');
formidable = require('formidable');
fs = require('fs');

// local server started at http://localhost:3000
app.listen(3000, (err)=> {console.log('connected!!')});

const input_script = require('./input-script');

// index.html
app.get('/', function(req, res){
  res.render('index', {filepath:''})
})

// download file
app.get('/download',function(req, res){
 
  res.download(__dirname+'/output.xlsx', 'user.xlsx')
})

// upload file post request
app.post('/fileupload', function(req, res){
  var form = new formidable.IncomingForm();
        form.parse(req, function (err, fields, files) {
            var prevpath = files.filename.path;
            var name = files.filename.name;
            var newpath = './' + name;

            var is = fs.createReadStream(prevpath);
            var os = fs.createWriteStream(newpath);

                // more robust than using fs.rename
                // allows moving across different mounted filesystems
            is.pipe(os);

                // removes the temporary file originally 
                // created by the post action
            is.on('end',function() {
                fs.unlinkSync(prevpath);
            });

             
            //calling read file function of input-script
            input_script.readFile();
            // storing the filepath of output file created 
            //sending the path  of output file on html page for downloading
            var filepathLink ='/output.xlsx';
            res.render('index', {filepath: 'Please download the file now'} )
        });
});