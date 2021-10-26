/**
 * Modules for creating http server
 */
 const express = require('express');
 const app = express();
 
 /**
  * Modules for working with the filesystem
  */
 const multer = require('multer');
 const fs = require('fs')
 const path = require('path')
 
 /**
  * Modules for parsing Excel files
  */
 const reader = require('xlsx')
 
 /**
  * Modules for generating table and
  * defining config. for the converted table
  */
 const jsonToTable = require('json-to-table');
 const table = require('table')
 const config = {
     border: table.getBorderCharacters('norc')
 };
 
 /**
  * Params for server and directory to save converted files
  */
 const hostname = '127.0.0.1';
 const port = 3010;
 const fileOutputDir = './output';
 
 app.use(express.static(path.join(__dirname, '/')));
 
 app.use(function(req, res, next) { //allow cross origin requests
     res.setHeader("Access-Control-Allow-Origin", "*");
     res.header("Access-Control-Allow-Methods", "POST, PUT, OPTIONS, DELETE, GET");
     res.header("Access-Control-Max-Age", "3600");
     res.header("Access-Control-Allow-Headers", "Content-Type, Access-Control-Allow-Headers, Authorization, X-Requested-With");
     next();
 });
 
 var storage = multer.diskStorage({ //multer's disk storage settings
     destination: function (req, file, cb) {
         cb(null, './uploads/')
     },
     filename: function (req, file, cb) {
         var datetimestamp = Date.now();
         cb(null, file.fieldname + '-' + datetimestamp + '.' + file.originalname.split('.')[file.originalname.split('.').length -1])
     }
 });
 
 var upload = multer({ //multer settings
                 storage: storage,
                 fileFilter : function(req, file, callback) { //file filter
                     if (['xls', 'xlsx'].indexOf(file.originalname.split('.')[file.originalname.split('.').length-1]) === -1) {
                         return callback(new Error('Wrong extension type!'));
                     }
                     callback(null, true);
                 }
             }).single('file');

 /** API path that will upload the files */
 app.post('/upload', function(req, res) {
 
     var exceltojson;
 
     upload(req,res,function(err){
         if(err){
             res.json({error_code:1,err_desc:err});
             return;
         }
 
         /** Check to see if user selected a file
          * (we should have file info in req.file object) */
         if(!req.file){
             res.json({error_code:1,err_desc:"No file passed"});
             return;
         }
 
 
         /**
          * Converting excel file to array of objects by using xlsxtojson module,
          * then processing it to .txt in form of human-readable table
          */
         try {
            /**
             * TODO:
             * Get custom sheets and columns for each sheet from user input on the front-end
             * (Nested check-box with possible values displayed)
             */

            const columnsInput= req.body.colrange.split(',');
            const sheetsInput = req.body.shrange.split(',');

            columnsInput.forEach((column) => {column.trim()});
            sheetsInput.forEach((sheet) => {sheet.trim()});

            const file = reader.readFile(req.file.path);
            const sheets = file.SheetNames;

            let resultText = "";

            for(let i = 0; i < sheets.length; i++)
            {

               let index = sheetsInput.indexOf(file.SheetNames[i]);


               if (index !== -1){
                    // let customColumnsForSheet = customSheetsAndColumnsInput[customSheetsInput[index]];
                    let sheet = file.Sheets[file.SheetNames[i]];
                    var temp = reader.utils.sheet_to_json(file.Sheets[file.SheetNames[i]], {header:"A"});
                    temp.forEach((element) => {
                        for (const key of Object.keys(element)){
                            if (columnsInput.indexOf(key) === -1){
                                delete element[key];
                            }
                        }
                    })

                    /**
                     * Remove 'A', 'B', ... from table header (use column original names)
                     */
                    temp = reader.utils.sheet_to_json(reader.utils.json_to_sheet(temp, {skipHeader:true}));

                    /**
                     * Convert json to string and append to .txt file in custom format 
                     */
                    resultText += `\nSHEET \'${file.SheetNames[i]}\':\n`;
                    resultText += table.table(jsonToTable(temp), config);
               }          
            }

            const fileName = req.file.originalname.substr(0, req.file.originalname.lastIndexOf('.'));
            const outputPath = path.resolve(__dirname, fileOutputDir, `${fileName}.txt`);

            const sheetsNum = sheetsInput.length;
            const colsNum = columnsInput.length;

            /**
             * Writing table to .txt file
             */
            fs.writeFile(outputPath, resultText, 'utf8', function (err) {
                if (err) {
                    return console.log(err);
                }
            
                console.log(`File succesfully converted to: ${outputPath}!\n
                Processed: ${sheetsNum} sheet(s) and ${colsNum} column(s).`);
            }); 

            res.write(`File succesfully converted to: ${outputPath}!\n
                        Processed: ${sheetsNum} sheet(s) and ${colsNum} column(s).`);
            res.send();

            /** Delete uploaded .xls/.xlsx file from uploads after conversion */
            try {
                fs.unlinkSync(req.file.path);
            } catch(e) {
                res.write('Error deleting the uploaded file!');
            }
            
         } catch (e){
             res.write('Corrupted excel file!');
         }
     })
 });
 
 app.get('/',function(req,res){
     res.sendFile(__dirname + "/index.html");
 });
 
 app.listen(port, function(){
     console.log(`Server running at http://${hostname}:${port}/`);    });