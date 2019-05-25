import * as XLSX from 'ts-xlsx';
import { Component } from '@angular/core';
import { ExcelServiceService } from'./excel-service.service';
import { Packer } from 'docx';
import { saveAs } from 'file-saver';
import * as XLSXG from 'xlsx';
import { DocumentCreator } from './cv-generator';
//interfaces
import { experiences, education, skills, achievements } from './cv-data';


@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {

    arrayBuffer:any;
    //new workbook instance
    wb = XLSXG.utils.book_new();

    file:File;

    constructor(private excelgenerator: ExcelServiceService)
    {

    }
    //array for the dummy table
    data: any = [{
	Num: '1',
	Name: 'JJ',
	SN: 'R',
	Age: 21
    },{
	Num: '2',
	Name: 'P',
	SN: 'B',
	Age: 27
    },{
	Num: '3',
	Name: 'T',
	SN: 'N',
	Age: 25
    }];

    //function to get full file's path
    public incomingfile(event):void
    {
	this.file= event.target.files[0];
    }

    //function to get excel data in json
   public upload():void {
	let fileReader = new FileReader();
        fileReader.onload = (e) => {
            this.arrayBuffer = fileReader.result;
            var data = new Uint8Array(this.arrayBuffer);
            var arr = new Array();
            for(var i = 0; i != data.length; ++i) arr[i] = String.fromCharCode(data[i]);
            var bstr = arr.join("");
            var workbook = XLSX.read(bstr, {type:"binary"});
            var first_sheet_name = workbook.SheetNames[0];
            var worksheet = workbook.Sheets[first_sheet_name];
            console.log(XLSX.utils.sheet_to_json(worksheet,{raw:true}));
        }
        fileReader.readAsArrayBuffer(this.file);
    }

    //function to export data as excel
    public exportAsXLSX():void {
	this.excelgenerator.exportAsExcelFile(this.data, 'sample');
    }

    //function to download generated word file
    public download(): void {
	const documentCreator = new DocumentCreator();
	const doc = documentCreator.create([experiences, education, skills, achievements]);

	const packer = new Packer();

	packer.toBlob(doc).then(blob => {
	    console.log(blob);
	    saveAs(blob, "example.docx");
	    console.log("Document created successfully");
	});
    }

     public create_excel_styles():void
    {

	 function bin_to_ex(s){

             var buf = new ArrayBuffer(s.length);
             var view = new Uint8Array(buf);
             for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
             return buf;

	 }

	 console.log('>>>>>PRINTING SOMETHING');
	 this.wb.Props = {
             Title: "SheetJS Tutorial",
             Subject: "Test",
             Author: "Red Stapler",
             CreatedDate: new Date(2017,12,19)
         };

	 this.wb.SheetNames.push("Test Sheet");
	 var ws_data = [['hello' , 'world']];  //a row with 2 columns
	 console.log('>>>>>> DATA',ws_data);
	 var ws = XLSX.utils.aoa_to_sheet(ws_data);
	 console.log('>>>>>> BINARYY FILEEEE',ws);
	 //this.wb.Sheets["Test Sheet"] = ws;
	 /* bookType can be 'xlsx' or 'xlsm' or 'xlsb' */
	 var wbout = XLSXG.write(this.wb, {bookType:'xlsx',  type: 'binary', cellStyles: true});
	 console.log(wbout);

	 /* the saveAs call downloads a file on the local machine */
	 saveAs(new Blob([bin_to_ex(wbout)],{type:""}), "excel_test.xlsx")
     }

}
