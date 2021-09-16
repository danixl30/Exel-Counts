const express = require('express');
const morgan = require('morgan');
const cors = require('cors');
const multer = require('multer');
const path = require('path');
const excelToJson = require('convert-excel-to-json');
const ExcelJS = require('exceljs');
const workbook = new ExcelJS.Workbook();
const fs = require('fs');
const { connect } = require('http2');

const app = express();

app.set('port', process.env.PORT || 8000);

app.use(morgan('dev'));
app.use(express.json());
app.use(cors());

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'views/index.html'));
})

const storage = multer.diskStorage({
    destination: path.join(__dirname, '/public'),
    filename: (req, file, cb) => {
        cb(null, file.originalname)
    }
});

app.post('/computeFiles', async (req, res) => {
    const upload =  multer({
        storage
    }).fields([{name: "baseFile"}, {name: "finalFile"}]);

    upload(req, res, async (err) => {
        if(!err){
            const baseFile = req.files.baseFile[0].filename;
            const finalFile = req.files.finalFile[0].filename;
            //console.log(req.files.finalFile[0].filename);
            var lines = [];
            const result = excelToJson({
                sourceFile: './public/' + baseFile
            });
            Object.keys(result).forEach(v => {
                Object.keys(result[v]).forEach(c => {
                    lines.push(result[v][c]);
                })
            })
            try{
                fs.unlinkSync(path.join(__dirname, '/public/file.xlsx'));
            }catch (e){
                console.log('File not found deleted');
            }
            //console.log(lines);
            let type = 1;
            const recordsMap = new Map();
            for (let i = 0; i < lines.length; i++){
                //console.log(lines[i]);
                if (lines[i].D && lines[i].D === 'TIPO DOC'){
                    type = 2;
                }
                if (lines[i].A && typeof lines[i].A === 'number'){
                    const id = lines[i].B;
                    if (recordsMap.has(id)){
                        const obj = recordsMap.get(id);
                        obj.amount += type === 1 ? lines[i].D : lines[i].E
                        recordsMap.set(id, obj);
                    }else{
                        const obj = {
                            id: lines[i].B,
                            reason: lines[i].C,
                            amount: type === 1 ? lines[i].D : lines[i].E
                        }
                        recordsMap.set(id, obj);
                    }
                }
            }
            //console.log(recordsMap.values());
            
            await workbook.xlsx.readFile(path.join(__dirname, '/public/'+finalFile))//Change file name here or give file path
            .then(async function() {
                var worksheet = workbook.getWorksheet('Hoja1');
                var i=1;
                var begininIndex;
                await worksheet.eachRow({ includeEmpty: false }, async function(row, rowNumber) {
                    r= await worksheet.getRow(i).values;
                    r1=r[2];// Indexing a column    
                    if (r[2] && r[2] === 'Nº' && r[3] && r[3] === 'RUT EMPRESA' && r[4] && r[4] === 'RAZON SOCIAL CLIENTE'){
                        begininIndex = i+1;
                        return;
                    }    
                    i++;
                }); 
                let num = 0;
                begininIndex = 13;
                for (const value of recordsMap.values()){    
                    let currentIndex = begininIndex+num;                  
                    let currentNum = num+1;
                    worksheet.getCell('B'+ currentIndex.toString()).value = currentNum;
                    worksheet.getCell('C'+ currentIndex.toString()).value = value.id;
                    worksheet.getCell('D'+ currentIndex.toString()).value = value.reason;
                    worksheet.getCell('E'+ currentIndex.toString()).value = value.amount;
                    num++;
                }
                ////Change the cell number here
                return workbook.xlsx.writeFile(path.join(__dirname, '/public/file.xlsx'))
                //Change file name here or give     file path
            });
            //res.json({state: true, msg: ""})
            await res.download(path.join(__dirname, '/public/file.xlsx'));
            //fs.unlinkSync(path.join(__dirname, '/public/file.xlsx'));
            fs.unlinkSync(path.join(__dirname, '/public/'+finalFile));
            fs.unlinkSync(path.join(__dirname, '/public/'+baseFile));
        }else{
            console.log(err);
            res.json({state: false, msg: "error"})
        }
    });
})

app.listen(app.get('port'), () => {
    console.log('Server started');
})