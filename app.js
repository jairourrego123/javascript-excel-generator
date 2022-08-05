

const ExcelJS = require('exceljs');
const expect = require('expect')

async function  GenerarReporte() {

    const workbook = new ExcelJS.Workbook();

    workbook.views = [ // controlan cuántas ventanas separadas Excel abrirá al ver el libro de trabajo.
        {
          x: 0, y: 0, width: 10000, height: 20000,
          firstSheet: 0, activeTab: 1, visibility: 'visible'
        }
      ]
      const sheet = workbook.addWorksheet('My Sheet', {properties:{tabColor:{argb:'0000000'}}}); //  agregar hoja de trabajo 
      const worksheet = workbook.getWorksheet('My Sheet')

      worksheet.mergeCells('A1', 'AS1'); // CINBINAR CELDAS DE CELDA A CELDA
      worksheet.mergeCells('A2', 'D2'); // CINBINAR CELDAS DE CELDA A CELDA
      worksheet.mergeCells('E2', 'I2'); // CINBINAR CELDAS DE CELDA A CELDA
      worksheet.mergeCells('J2', 'N2'); // CINBINAR CELDAS DE CELDA A CELDA
      worksheet.mergeCells('O2', 'P2'); // CINBINAR CELDAS DE CELDA A CELDA
      worksheet.mergeCells('Q2', 'R2'); // CINBINAR CELDAS DE CELDA A CELDA
      worksheet.mergeCells('S2', 'T2'); // CINBINAR CELDAS DE CELDA A CELDA
      worksheet.mergeCells('W2', 'Z2'); // CINBINAR CELDAS DE CELDA A CELDA
      worksheet.mergeCells('AA2', 'AE2'); // CINBINAR CELDAS DE CELDA A CELDA
      worksheet.mergeCells('AF2', 'AJ2'); // CINBINAR CELDAS DE CELDA A CELDA
      worksheet.mergeCells('AK2', 'AL2'); // CINBINAR CELDAS DE CELDA A CELDA
      worksheet.mergeCells('AM2', 'AR2'); // CINBINAR CELDAS DE CELDA A CELDA

      // VALORES FILA 2


//     //   worksheet.columns = [
//     //     { header: 'Id', key: 'id', width: 10 },
//     //     { header: 'Name', key: 'name', width: 32,style: { font: { color: { argb: 'FF00FF00' }} }},
//     //     { header: 'D.O.B.', key: 'dob', width: 40 }
//     //   ];

 
//    //worksheet.getRow(2).style = { font:font , fill:backgroundRow2}
// //    worksheet.getRow(2).fill(backgroundRow2).font(font)

   
  


 

   
      worksheet.getCell('A1').value = 'REGISTRO DE AUDIENCIAS DE CONCILIACIÓN';
      worksheet.getCell('A2').value = 'DATOS DE SOLICITUD DE AUDIENCIA';
      worksheet.getCell('E2').value = 'DATOS DE SOLICITUD DE AUDIENCIA';
      worksheet.getCell('J2').value = 'DATOS CONVOCADO';
      worksheet.getCell('O2').value = 'FECHA AUDIENCIA';
      worksheet.getCell('Q2').value = 'ESTADO';
      worksheet.getCell('S2').value = 'RESULTADO DEL TRÁMITE';
      worksheet.getCell('U2').value = 'SEGUIMIENTO';
      worksheet.getCell('V2').value = 'SNIES';
      worksheet.getCell('W2').value = '';
      worksheet.getCell('AA2').value = 'EVALUACIÓN DEL CONCILIADOR';
      worksheet.getCell('AF2').value = 'EVALUACIÓN DEL CENTRO';
      worksheet.getCell('AK2').value = 'EVALUACIÓN DEL MECANISMO';
      worksheet.getCell('AM2').value = '¿POR CUÁL MEDIO CONOCIÓ EL CENTRO DE CONCILIACION? ';

      //  Style 
      const font = {name: 'Calibri', family: 4, size: 10,color:{argb:'FFFFFF'} }; // 
      const backgroundRow2 = {type: 'pattern',pattern:'solid',fgColor:{argb:'A0A4FA'},bgColor:{argb:'A0A4FA'}};
      const backgroundRow3 = {type: 'pattern',pattern:'solid',fgColor:{argb:'7ABF96'},bgColor:{argb:'7ABF96'}};

      const border  = {top: {style:'thin'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};
   
      worksheet.getRows(1).font=font
           // Ajuste Filas

      worksheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };
      worksheet.getRow(2).alignment = { vertical: 'middle', horizontal: 'center' };
      worksheet.getRow(2).fill= backgroundRow2
       
      worksheet.getRow(2).font= font 
      worksheet.getRow(2).height= 30;
      



        worksheet.columns = [
                { header: 'FECHA SOLICITUD', key: 'Fecha_registro',width:20},
                { header: 'No. TRAMITE', key: 'Numero_caso', width: 32},// { header: 'No. De Tramite', key: 'numero_solicitud', width: 32,style: { font: { color: { argb: 'FF00FF00' }} }},
                { header: 'MATERIA', key: 'Area_Id', width: 40 },
                { header: 'ASUNTO', key: 'Tema', width: 40 },
                { header: 'CONVOCANTE', key: 'convocante', width: 40 },
                { header: 'NO DE DOCUMENTO', key: 'documento_convocante', width: 40 },
                { header: 'GENERO', key: 'genero_convocante', width: 40 },
                { header: 'ESTRATO', key: 'estrato_convocante', width: 40 },
                { header: 'LOCALIDAD', key: 'localidad_convocante', width: 40 },
                { header: 'CONVOCADO', key: 'convocado', width: 40 },
                { header: 'NO DE DOCUMENTO', key: 'documento_convocado', width: 40 },
                { header: 'GENERO', key: 'genero_convocado', width: 40 },
                { header: 'ESTRATO', key: 'estrato_convocado', width: 40 },
                { header: 'LOCALIDAD', key: 'localidad_convocado', width: 40 },
                { header: 'MES', key: 'mes', width: 40 },
                { header: 'FECHA  DE AUDIENCIA', key: 'fecha_audiencia', width: 40 },
                { header: 'ESTADO DEL TRAMITE', key: 'estado_solicitud', width: 40 },
                { header: 'NUEVA FECHA', key: 'nueva_fecha', width: 40 },
                { header: 'RESULTADO DEL TRÁMITE ', key: 'resultado', width: 40 },
                { header: 'No. RESULTAD', key: 'numero_resultado', width: 40 },
                { header: 'CUMPLIO', key: 'cumplio', width: 40 },
                { header: 'POBLACIÓN CICLO VITAL', key: 'poblacion', width: 40 },
                { header: 'CONCILIADOR', key: 'conciliador', width: 40 },
                { header: 'RUG', key: 'rug', width: 40 },
                { header: 'COMISARIA', key: 'comisaria', width: 40 },
                { header: 'REMITE', key: 'remite', width: 40 },
                { header: 'Servicio recibido por parte del conciliador', key: 'pregunta1', width: 40 },
                { header: 'Puntualidad', key: 'pregunta2', width: 40 },
                { header: 'Dominio del tema', key: 'pregunta3', width: 40 },
                { header: 'Lenguaje utilizado', key: 'pregunta4', width: 40 },
                { header: 'Manejo de la audiencia', key: 'pregunta5', width: 40 },
                { header: 'Servicio prestado por el centro', key: 'pregunta6', width: 40 },
                { header: 'Imparcialidad', key: 'pregunta7', width: 40 },
                { header: 'Satisfacción por la información que brindada el ', key: 'pregunta8', width: 40 },
                { header: 'Satisfacción por el tiempo de atención', key: 'pregunta9', width: 40 },
                { header: 'Amabilidad del personal del Centro', key: 'pregunta10', width: 40 },
                { header: '¿La conciliación lleno sus expectativas en el tratamiento del ', key: 'pregunta11', width: 40 },
                { header: '¿Recomendaría la conciliación para resolver ', key: 'pregunta12', width: 40 },
                { header: 'Radio', key: 'pregunta13', width: 40 },
                { header: 'Folletos', key: 'pregunta14', width: 40 },
                { header: 'Televisión', key: 'pregunta15', width: 40 },
                { header: 'Un amigo', key: 'pregunta16', width: 40 },
                { header: 'Web', key: 'pregunta17', width: 40 },
                { header: 'Otro', key: 'pregunta18', width: 40 },
                { header: 'FECHA DE ENTREGA', key: 'fecha_entrega', width: 40 },

              ];

             
            const rows = {}
            for (const iterator of worksheet.columns) {
                rows[iterator.key]=iterator.header
            }

            worksheet.addRow(rows, 'n');

            worksheet.getRow(3).alignment = { vertical: 'middle', horizontal: 'center' };
            worksheet.getRow(3).height= 40;
            worksheet.getRow(3).fill= backgroundRow3
            worksheet.getRow(3).border=border



            datos=[{Fecha_registro:2200,Numero_caso:20,numero_solicitud:1},{fecha_solicitud:11,Numero_caso:2},{fecha_solicitud:210,numero_solicitud:5,Numero_caso:8}]
            for (const iterator of datos) {
                worksheet.addRow(iterator, 'n').border=border; 
            }
             

    //   const idCol = worksheet.getColumn('id');
    //   const nameCol = worksheet.getColumn('B');
    //   const dobCol = worksheet.getColumn(3);


      

        
    //   // set column properties

    // // Note: will overwrite cell value C1
    // dobCol.header = 'Date of Birth';

    // // Note: this will overwrite cell values C1:C2
    // dobCol.header = ['Date of Birth'];


    // worksheet.addRow({id: 1, name: 'John Doe', dob: new Date(1970,1,1)});
    // worksheet.addRow({id: 2, name: 'Jane Doe', dob: new Date(1965,1,7)})
    // worksheet.addRow({id: 3, name: 'Jane Doe', dob: new Date(1965,1,7)});

  
      //worksheet.mergeCells('A4:B5');
      //worksheet.getCell('B5').style.font = myFonts.arial;
      //worksheet.getCell('B5').value = 'Hello, World!';
      
     


      //worksheet.mergeCells(10,11,12,13);

      // Insert a couple of Rows by key-value, shifting down rows every time



    // worksheet.getCell('A4').numFmt = '0.00%'; // estilo a una celda 



      await workbook.xlsx.writeFile("Formato Consolidado.xlsx")
   
}
  GenerarReporte();


// const XLSX = require('xlsx');

// const XlsxTemplate = require('xlsx-template');


// const convertJsonToExcel=()=>{
    
//     const workSheet=XLSX.utils.json_to_sheet(student) // crear hoja
//     XLSX.utils.sheet_add_aoa(workSheet,[["nombre","edad"]],{origin:"A1"})
//    // const workBook=XLSX.utils.book_new(); // crear libro

//     XLSX.utils.book_append_sheet(workBook,workSheet,"STUDENTS")
//     // generar bufer 
//     XLSX.write(workBook,{bookType:'xlsx',type:"buffer"})

//     // binaria string

//     XLSX.write(workBook,{bookType:"xls",type:"binary"})

//     XLSX.writeFile(workBook,"estudiantesDatos.xlsx") // descrgar el archivo con el nomnbre estudiantes Datos
// }

// convertJsonToExcel()