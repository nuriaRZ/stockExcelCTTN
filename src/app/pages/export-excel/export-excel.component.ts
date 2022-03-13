import { Component, OnInit } from '@angular/core';
import { NgxFileDropEntry, FileSystemFileEntry, FileSystemDirectoryEntry } from 'ngx-file-drop';
import * as XLSX from 'xlsx';
import * as FileSaver from 'file-saver';
import { FormControl, FormGroup, Validators } from '@angular/forms';

@Component({
  selector: 'app-export-excel',
  templateUrl: './export-excel.component.html',
  styleUrls: ['./export-excel.component.scss']
})
export class ExportExcelComponent implements OnInit {
  form = new FormGroup({
    
    fileField: new FormControl('', [Validators.required]),
    
  });
  file: any;
  arrayBuffer: any;
  fileList: any;
  data: any = [];
  colorsCells = ['FFD9D9D9', 'FFBFBFBF', 'FFA6A6A6', 'FF919191', 'FF808080', 'FF757575', 'FF686868', 'FF585858', 'FF525252', 'FF4C4C4C' ]
  

  constructor() { }

  ngOnInit(): void {
  }

  upload(event){
    {    
      this.file= event.target.files[0];     
      let fileReader = new FileReader();    
      fileReader.readAsArrayBuffer(this.file);     
      fileReader.onload = (e) => {    
          this.arrayBuffer = fileReader.result;    
          var data = new Uint8Array(this.arrayBuffer);    
          var arr = new Array();    
          for(var i = 0; i != data.length; ++i) arr[i] = String.fromCharCode(data[i]);    
          var bstr = arr.join("");    
          var workbook = XLSX.read(bstr, {type:"binary"});    
          var first_sheet_name = workbook.SheetNames[0];    
          var worksheet = workbook.Sheets[first_sheet_name];    
          // console.log(XLSX.utils.sheet_to_json(worksheet,{raw:true}));    
            var arraylist = XLSX.utils.sheet_to_json(worksheet,{raw:true});     
                this.fileList = [];    
                this.fileList = arraylist;
                console.log(this.fileList)    
          this.mappingDataSageExcel(this.fileList);
        
      } 
    }
  }


  mappingDataSageExcel(sageExcel){
    let regex = new RegExp(/-/);
    for (let i = 11; i < sageExcel.length; i++) {
      const element = sageExcel[i];
      let ref =  element['STOCKS DE ARTICULOS POR ALMACEN DESGLOSADO'].toString();
      let artRef;
      let artColor;     
      // if(ref.match(regex) != null){
      //   let separatorIndex = ref.match(regex).index;
      //   artRef = ref.substr(0, separatorIndex);
      //   artColor = ref.substr(separatorIndex + 1, ref.length);
      //   console.log('lasize',artColor);
      // }else {
      //   artRef = ref
      // }
      
      if(ref !== '*'){
        if(this.data.length == 0){
          let articulo = {
            referencia: ref,
            descripcion: element['__EMPTY'],
            talla: [{index:element['__EMPTY_1'], value: 1}],
            color: '',
          }
          this.data.push(articulo);
        }else{
          if(this.data.length > 0){
            let artFound;
            artFound = this.data.find(d => d.referencia == ref)
            console.log('foundedArt', artFound);
            if(artFound != undefined){
              
              artFound.talla.forEach(size => {
                if(size.index == element['__EMPTY_1']){
                  size.value = size.value + 1;  
                }else{
                  artFound.talla.push( {
                    index : element['__EMPTY_1'],
                    value : 1,
                  })
                }
              })
              
    
            }else{
              let articulo = {
                referencia: ref,
                descripcion: element['__EMPTY'],
                talla: [{index:element['__EMPTY_1'], value: 1}],
                color: '',
              }
              this.data.push(articulo);
            }
          }
        }
        
        
      }
    }
    console.log('data', this.data);
    this.convertToExcelFile(this.data)
  }

  convertToExcelFile(dataExcel){
   
    const Excel = require('exceljs');

    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet("My Sheet");
    workbook.views = [
      {
        x: 0, y: 0, width: 10000, height: 20000,
        firstSheet: 0, activeTab: 1, visibility: 'visible'
      }
    ];
    worksheet.columns = [{ width: 11 }, { width: 25 }]

    
    
    const title = 'Titulo';
    

   
    let trajes = [];
    let camisas = [];
    let chaquetas = [];
    let pantalones = [];
    let tejanos = [];
    let jerseys = [];
    let cazadoras = [];
    let polos = [];
    let baño = [];
    let chalecos = [];
    let bermudas = [];
    let calzado = [];
    let sudaderas = [];
    let cinturones = [];
    let corbatas = [];
    let complementos = [];
    let carteras = [];

    dataExcel.forEach(element => {
      let c = element.referencia.toString();
      let code = c.substr(0,2);
      console.log(code);
      switch (code) {
        case '01':
          trajes.push(element);  
          break;

        case '02':
          chaquetas.push(element);
          break;
        case '03':
          pantalones.push(element);  
          break;
        case '04':
          camisas.push(element);  
          break;
        case '05':
          jerseys.push(element);  
          break;        
        case '07':
          sudaderas.push(element);  
          break;
        case '08':
          cazadoras.push(element);  
          break; 
        case '15':
          polos.push(element);  
          break; 
        case '17':
          tejanos.push(element);  
          break;
        case '18':
          chalecos.push(element);  
          break;
        case '19':
          bermudas.push(element);  
          break;
        case '20':
          cinturones.push(element);  
          break;
        case '21':
          corbatas.push(element);  
          break;
        case '22':
          carteras.push(element);  
          break;
        case '23':
          complementos.push(element);  
          break;
        case '24':
          calzado.push(element);  
          break;
        default:
          break;
      }
      
      if(element.descripcion.includes('Tejano')){
        tejanos.push(element);        
      }
      
    });
    let tallaColumn = worksheet.getColumn(4);
   let trajeRow = worksheet.addRow(['REF', 'TRAJES', '44', '46', '48', '50', '52', '54', '56', '58' ]);
     trajeRow.font = { name: 'Arial', family: 4, size: 10, bold: true, color: {argb: "ffffff"} };
     trajeRow.fill = {
      type: 'pattern',
      pattern:'solid',
      fgColor:{ argb:'000080' }
      };

      
      for (let i = 0; i < trajes.length; i++) {
        const element = trajes[i];
       let trajesRow = worksheet.addRow([element.referencia, element.descripcion,  element.color, 0, 0, 0, 0, 0, 0, 0]);
        trajeRow.eachCell(cell =>{
          if(cell['_column']['_number'] > 2){
            cell.alignment = {
              vertical: 'middle',
              horizontal: 'center'
            }
          }
          let indexColor = 0;
            trajesRow.eachCell(data =>{
             this.addStylesInCell(data, indexColor);
                  element.talla.forEach(size => {
                    if(cell.value == size.index && cell['_column']['_number'] == data['_column']['_number']){
                      data.value = size.value                      
                  }
              })
              indexColor++;
                
            })
        })
          
      }   
        
  
    
   let chaqueRow = worksheet.addRow(['REF', 'CHAQUETAS', '46', '48', '50', '52', '54', '56', '58' ]);
   chaqueRow.font = { name: 'Arial', family: 4, size: 10, bold: true, color: {argb: "ffffff"} };
   chaqueRow.fill = {
      type: 'pattern',
      pattern:'solid',
      fgColor:{ argb:'000080' }
      };

      for (let i = 0; i < chaquetas.length; i++) {
        const element = chaquetas[i];
        let chaquetasRow = worksheet.addRow([element.referencia, element.descripcion,  element.color, 0, 0, 0, 0, 0, 0]);
        chaqueRow.eachCell(cell =>{
          if(cell['_column']['_number'] > 2){
            cell.alignment = {
              vertical: 'middle',
              horizontal: 'center'
            }
          }
          let indexColor = 0;
          chaquetasRow.eachCell(data =>{
            this.addStylesInCell(data, indexColor);
                  element.talla.forEach(size => {
                    if(cell.value == size.index && cell['_column']['_number'] == data['_column']['_number']){
                      data.value = size.value                      
                    }
                  })
            indexColor++;      
                
          })
        })
      }  
    

    
    let pantRow = worksheet.addRow(['REF', 'PANTALONES', '38', '40', '42', '44', '46', '48', '50' ]);
    pantRow.font = { name: 'Arial', family: 4, size: 10, bold: true, color: {argb: "ffffff"}, align: 'center', };
    pantRow.fill = {
        type: 'pattern',
        pattern:'solid',
        fgColor:{ argb:'000080' }
        };
        for (let i = 0; i < pantalones.length; i++) {
          const element = pantalones[i];
          let pantalonesRow = worksheet.addRow([element.referencia, element.descripcion,  element.color, 0, 0, 0, 0, 0, 0]);
          pantRow.eachCell(cell =>{   
            let indexColor = 0;
            pantalonesRow.eachCell(data =>{
              if(cell['_column']['_number'] > 2){
                cell.alignment = {
                  vertical: 'middle',
                  horizontal: 'center'
                }
              }
              this.addStylesInCell(data, indexColor);
                element.talla.forEach(size => {
                  if(cell.value == size.index && cell['_column']['_number'] == data['_column']['_number']){
                    data.value = size.value                      
                  }
                })
                indexColor++;    
            })
          })
        }
    
    let tejRow = worksheet.addRow(['REF', 'TEJANOS', '38', '40', '42', '44', '46', '48', '50' ]);
    tejRow.font = { name: 'Arial', family: 4, size: 10, bold: true, color: {argb: "ffffff"} };
    tejRow.fill = {
        type: 'pattern',
        pattern:'solid',
        fgColor:{ argb:'000080' }
        };
        for (let i = 0; i < tejanos.length; i++) {
          const element = tejanos[i];
          let tejanosRow = worksheet.addRow([element.referencia, element.descripcion,  element.color, 0, 0, 0, 0, 0, 0]);
          tejRow.eachCell(cell =>{
            let indexColor = 0;
            tejanosRow.eachCell(data =>{
              if(cell['_column']['_number'] > 2){
                cell.alignment = {
                  vertical: 'middle',
                  horizontal: 'center'
                }
              }
              this.addStylesInCell(data, indexColor);
                element.talla.forEach(size => {
                  if(cell.value == size.index && cell['_column']['_number'] == data['_column']['_number']){
                    data.value = size.value                      
                  }
                })
                indexColor++;  
            })
          })
        }

    let camisasRow = worksheet.addRow(['REF', 'CAMISAS', '38', '40', '42', '44', '46', '48', '50' ]);
    camisasRow.font = { name: 'Arial', family: 4, size: 10, bold: true, color: {argb: "ffffff"} };
    camisasRow.fill = {
        type: 'pattern',
        pattern:'solid',
        fgColor:{ argb:'000080' }
        };
        for (let i = 0; i < camisas.length; i++) {
          const element = camisas[i];
          let camiRow = worksheet.addRow([element.referencia, element.descripcion,  element.color, 0, 0, 0, 0, 0, 0]);
          camisasRow.eachCell(cell =>{
            let indexColor = 0;
            camiRow.eachCell(data =>{
              if(cell['_column']['_number'] > 2){
                cell.alignment = {
                  vertical: 'middle',
                  horizontal: 'center'
                }
              }
              this.addStylesInCell(data, indexColor);
                element.talla.forEach(size => {
                  if(cell.value == size.index && cell['_column']['_number'] == data['_column']['_number']){
                    data.value = size.value                      
                  }
                })
                indexColor++;  
            })
          })
        }

    let jerRow = worksheet.addRow(['REF', 'JERSEYS', 'S', 'M', 'L', 'XL', 'XXL', '3XL', '' ]);
    jerRow.font = { name: 'Arial', family: 4, size: 10, bold: true, color: {argb: "ffffff"} };
    jerRow.fill = {
        type: 'pattern',
        pattern:'solid',
        fgColor:{ argb:'000080' }
        };
    for (let i = 0; i < jerseys.length; i++) {
          const element = jerseys[i];
          let jerseysRow = worksheet.addRow([element.referencia, element.descripcion,  element.color, 0, 0, 0, 0, 0,]);
          jerRow.eachCell(cell =>{
            let indexColor = 0;
            jerseysRow.eachCell(data =>{
              if(cell['_column']['_number'] > 2){
                cell.alignment = {
                  vertical: 'middle',
                  horizontal: 'center'
                }
              }
              this.addStylesInCell(data, indexColor);
                element.talla.forEach(size => {
                  if(cell.value == size.index && cell['_column']['_number'] == data['_column']['_number']){
                    data.value = size.value                      
                  }
                })
                indexColor++;                  
            })
          })
        }

    let cazRow = worksheet.addRow(['REF', 'CAZADORAS', 'S', 'M', 'L', 'XL', 'XXL', '3XL', '' ]);
    cazRow.font = { name: 'Arial', family: 4, size: 10, bold: true, color: {argb: "ffffff"} };
    cazRow.fill = {
        type: 'pattern',
        pattern:'solid',
        fgColor:{ argb:'000080' }
        };
     for (let i = 0; i < cazadoras.length; i++) {
          const element = cazadoras[i];
          let cazadorasRow = worksheet.addRow([element.referencia, element.descripcion,  element.color, 0, 0, 0, 0, 0]);
          cazRow.eachCell(cell =>{
            let indexColor = 0;
            cazadorasRow.eachCell(data =>{
              if(cell['_column']['_number'] > 2){
                cell.alignment = {
                  vertical: 'middle',
                  horizontal: 'center'
                }
              }
              this.addStylesInCell(data, indexColor);
                element.talla.forEach(size => {
                  if(cell.value == size.index && cell['_column']['_number'] == data['_column']['_number']){
                    data.value = size.value                      
                  }
                })
                indexColor++; 
            })  
          })
        }
    let polosRow = worksheet.addRow(['REF', 'POLOS', 'S', 'M', 'L', 'XL', 'XXL', '3XL', '' ]);
    polosRow.font = { name: 'Arial', family: 4, size: 10, bold: true, color: {argb: "ffffff"} };
    polosRow.fill = {
        type: 'pattern',
        pattern:'solid',
        fgColor:{ argb:'000080' }
        };
        for (let i = 0; i < polos.length; i++) {
          const element = polos[i];
          let poloRow = worksheet.addRow([element.referencia, element.descripcion,  element.color, 0, 0, 0, 0, 0]);
          polosRow.eachCell(cell =>{
            let indexColor = 0;
            poloRow.eachCell(data =>{
              if(cell['_column']['_number'] > 2){
                cell.alignment = {
                  vertical: 'middle',
                  horizontal: 'center'
                }
              }
              this.addStylesInCell(data, indexColor);
                element.talla.forEach(size => {
                  if(cell.value == size.index && cell['_column']['_number'] == data['_column']['_number']){
                    data.value = size.value                      
                  }
                })
                   indexColor++;  
            })
          })
        }
    let chalRow = worksheet.addRow(['REF', 'CHALECOS', 'S', 'M', 'L', 'XL', 'XXL', '3XL', '' ]);
    chalRow.font = { name: 'Arial', family: 4, size: 10, bold: true, color: {argb: "ffffff"} };
    chalRow.fill = {
        type: 'pattern',
        pattern:'solid',
        fgColor:{ argb:'000080' }
        };
        for (let i = 0; i < chalecos.length; i++) {
          const element = chalecos[i];
          let chalecosRow = worksheet.addRow([element.referencia, element.descripcion,  element.color, 0, 0, 0, 0, 0]);
          chalRow.eachCell(cell =>{
            let indexColor = 0;
            chalecosRow.eachCell(data =>{
              if(cell['_column']['_number'] > 2){
                cell.alignment = {
                  vertical: 'middle',
                  horizontal: 'center'
                }
              }
              this.addStylesInCell(data, indexColor);
                element.talla.forEach(size => {
                  if(cell.value == size.index && cell['_column']['_number'] == data['_column']['_number']){
                    data.value = size.value                      
                  }
                })
                   indexColor++;  
            })
          })
        }
    let bermRow = worksheet.addRow(['REF', 'BERMUDAS', '38', '40', '42', '44', '46', '48', '50' ]);
    bermRow.font = { name: 'Arial', family: 4, size: 10, bold: true, color: {argb: "ffffff"} };
    bermRow.fill = {
        type: 'pattern',
        pattern:'solid',
        fgColor:{ argb:'000080' }
        };
        for (let i = 0; i < bermudas.length; i++) {
          const element = bermudas[i];
          let bermudasRow = worksheet.addRow([element.referencia, element.descripcion,  element.color, 0, 0, 0, 0, 0, 0]);
          bermRow.eachCell(cell =>{
            let indexColor = 0;
            bermudasRow.eachCell(data =>{
              if(cell['_column']['_number'] > 2){
                cell.alignment = {
                  vertical: 'middle',
                  horizontal: 'center'
                }
              }
              this.addStylesInCell(data, indexColor);
                element.talla.forEach(size => {
                  if(cell.value == size.index && cell['_column']['_number'] == data['_column']['_number']){
                    data.value = size.value                      
                  }
                })
                   indexColor++;  
            })
          })
        }
    let cintoRow = worksheet.addRow(['REF', 'CINTURONES', '90', '95', '100', '105', '115', ]);
    cintoRow.font = { name: 'Arial', family: 4, size: 10, bold: true, color: {argb: "ffffff"} };
    cintoRow.fill = {
        type: 'pattern',
        pattern:'solid',
        fgColor:{ argb:'000080' }
        };
        for (let i = 0; i < cinturones.length; i++) {
          const element = cinturones[i];
          let cinturonesRow = worksheet.addRow([element.referencia, element.descripcion,  element.color, 0, 0, 0, 0]);
          cintoRow.eachCell(cell =>{
            let indexColor = 0;
            cinturonesRow.eachCell(data =>{
              if(cell['_column']['_number'] > 2){
                cell.alignment = {
                  vertical: 'middle',
                  horizontal: 'center'
                }
              }
              this.addStylesInCell(data, indexColor);
                element.talla.forEach(size => {
                  if(cell.value == size.index && cell['_column']['_number'] == data['_column']['_number']){
                    data.value = size.value                      
                  }
                })
                   indexColor++;  
            })
          })
        }
    let corbRow = worksheet.addRow(['REF', 'CORBATAS',  ]);
    corbRow.font = { name: 'Arial', family: 4, size: 10, bold: true, color: {argb: "ffffff"} };
    corbRow.fill = {
        type: 'pattern',
        pattern:'solid',
        fgColor:{ argb:'000080' }
        };
        for (let i = 0; i < corbatas.length; i++) {
          const element = corbatas[i];
          let corbatasRow = worksheet.addRow([element.referencia, element.descripcion,  element.color, 0, 0, 0, 0, 0, 0]);
          corbRow.eachCell(cell =>{
            let indexColor = 0;
            corbatasRow.eachCell(data =>{
              if(cell['_column']['_number'] > 2){
                cell.alignment = {
                  vertical: 'middle',
                  horizontal: 'center'
                }
              }
              this.addStylesInCell(data, indexColor);
                element.talla.forEach(size => {
                  if(cell.value == size.index && cell['_column']['_number'] == data['_column']['_number']){
                    data.value = size.value                      
                  }
                })
                   indexColor++;  
            })
          })
        }
    let cartRow = worksheet.addRow(['REF', 'CARTERAS', '38', '40', '42', '44', '46', '48', '50' ]);
    cartRow.font = { name: 'Arial', family: 4, size: 10, bold: true, color: {argb: "ffffff"} };
    cartRow.fill = {
        type: 'pattern',
        pattern:'solid',
        fgColor:{ argb:'000080' }
        };
        for (let i = 0; i < carteras.length; i++) {
          const element = carteras[i];
          let carterasRow = worksheet.addRow([element.referencia, element.descripcion,  element.color, 0, 0, 0, 0, 0, 0]);
          cartRow.eachCell(cell =>{
            let indexColor = 0;
            carterasRow.eachCell(data =>{
              if(cell['_column']['_number'] > 2){
                cell.alignment = {
                  vertical: 'middle',
                  horizontal: 'center'
                }
              }
              this.addStylesInCell(data, indexColor);
                element.talla.forEach(size => {
                  if(cell.value == size.index && cell['_column']['_number'] == data['_column']['_number']){
                    data.value = size.value                      
                  }
                })
                   indexColor++;  
            })
          })
        }
    let complRow = worksheet.addRow(['REF', 'COMPLEMENTOS', '38', '40', '42', '44', '46', '48', '50' ]);
    complRow.font = { name: 'Arial', family: 4, size: 10, bold: true, color: {argb: "ffffff"} };
    complRow.fill = {
        type: 'pattern',
        pattern:'solid',
        fgColor:{ argb:'000080' }
        };
        for (let i = 0; i < complementos.length; i++) {
          const element = complementos[i];
          let complementosRow = worksheet.addRow([element.referencia, element.descripcion,  element.color, 0, 0, 0, 0, 0, 0]);
          complRow.eachCell(cell =>{
            let indexColor = 0;
            complementosRow.eachCell(data =>{
              if(cell['_column']['_number'] > 2){
                cell.alignment = {
                  vertical: 'middle',
                  horizontal: 'center'
                }
              }
              this.addStylesInCell(data, indexColor);
                element.talla.forEach(size => {
                  if(cell.value == size.index && cell['_column']['_number'] == data['_column']['_number']){
                    data.value = size.value                      
                  }
                })
                   indexColor++;  
            })
          })
        }
    let calzRow = worksheet.addRow(['REF', 'CALZADO', '39', '40', '41', '42', '43', '44', '45' ]);
    calzRow.font = { name: 'Arial', family: 4, size: 10, bold: true, color: {argb: "ffffff"} };
    calzRow.fill = {
        type: 'pattern',
        pattern:'solid',
        fgColor:{ argb:'000080' }
        };
        for (let i = 0; i < calzado.length; i++) {
          const element = calzado[i];
          let calzadosRow = worksheet.addRow([element.referencia, element.descripcion,  element.color, 0, 0, 0, 0, 0, 0]);
          calzRow.eachCell(cell =>{
            let indexColor = 0;
            calzadosRow.eachCell(data =>{              
              if(cell['_column']['_number'] > 2){
                cell.alignment = {
                  vertical: 'middle',
                  horizontal: 'center'
                }
              }              
              this.addStylesInCell(data, indexColor);
                element.talla.forEach(size => {
                  if(cell.value == size.index && cell['_column']['_number'] == data['_column']['_number']){
                    data.value = size.value                      
                  }
                })
                   indexColor++;  
            })
          })
        }

    // worksheet.insertRows(1, trajes);

     workbook.xlsx.writeBuffer().then((data) => {
      let blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      FileSaver.saveAs(blob, 'inventarioCotton.xlsx');
      this.form.controls['fileField'].reset();
      this.file = '';
    });
  }

  paintCell(cell: any, color){
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: {
        argb: color
      }, 
    }
    cell.border = {
      top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" }
    }
  }

  addStylesInCell(cell: any, color){
    if(cell['_column']['_number'] > 2){
      cell.alignment = {
        vertical: 'middle',
        horizontal: 'center'
      }
     this.paintCell(cell, this.colorsCells[color]);
    }else{
      cell.border = {
        top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" }
      }
    }
  }

  saveAsExcelFile(buffer: any, fileName: string): void {
    const EXCEL_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
    const EXCEL_EXTENSION = '.xlsx';
    const data: Blob = new Blob([buffer], {type: EXCEL_TYPE});
    FileSaver.saveAs(data, fileName + '_export_' + new  Date().getTime() + EXCEL_EXTENSION);

 }
}

  

  
