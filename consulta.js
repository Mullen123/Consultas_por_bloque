
//npm install oracledb
//npm install prompt-sync
const oracledb = require('oracledb');
var moment = require('moment'); 
prompt = require("prompt-sync")({sigint :true});
var xl = require('excel4node');
// Create a new instance of a Workbook class
var wb = new xl.Workbook();
const path = require('path');

async function run() {
  var connection;

  var aux = new Array();
  var isr;
  var query_rn;
  var aux_rn;
  var query_rfc;
  var aux_rfc;
  var query_date;
  var aux_date;
  var query_total;
  var aux_total;
  let fecha;
  let month;
  let day;  
  let year;
  var first_date;
  var exp_yyyymmdd = /(\d{4})[./-](\d{2})[./-](\d{2})$/;
  var date2;


  /*funcion que te retornara  un array con todos los esquemas a consultar en la bse de datos ejemplo: isr,isr2.... isrn*/
  function datos(){

let esquema  = [''];

let aux_esquema = esquema.map(element=>{
  return element.split(" ");

});
let esquemas = JSON.stringify(aux_esquema);
let aux_isr  = esquemas.replace(/"/gi,"").trim().slice(2).slice(0, -2);
 aux = aux_isr.split(",");

  return aux;
}

try {

//generamos una conexión a la base de datos//
connection = await oracledb.getConnection({user: "", password: "", connectionString: ""});
console.log("Conexión Exitosa !!!!");


async function alldata(){
//asiganmos a una variable lo que nos retorna la funcion de datos ['isr2','isr10']//
isr = datos();


for (let i = 0 ;i<isr.length;i++){
//---------------------------  caso nomina y tipo regimen -------------------------//
query_rn  = await connection.execute(' select distinct(TIPONOMINA),TIPOREGIMEN from M4T_NOMINA12_CNR_33_'+ isr[i]); 
aux_rn = Object.values(query_rn.rows);

//--------------------------- rfc Emisor -------------------------//
query_rfc = await connection.execute(`SELECT DISTINCT(RFCEMISOR) FROM  M4T_COMPROBANTES_DC_33_${isr[i]} `); 
aux_rfc = Object.values(query_rfc.rows);


//---------------------------     total de registros -------------------------//
query_total = await connection.execute(`SELECT COUNT(RFCEMISOR) FROM  M4T_COMPROBANTES_DC_33_${isr[i]}`); 
aux_total = Object.values(query_total.rows);


//---------------------------  caso fecha -------------------------//
query_date =  await connection.execute('select distinct (FECHAPAGO) from M4T_NOMINA12_CNR_33_'+ isr[i]);
aux_date = Object.values(query_date.rows);

date_format();

isr[i] =  data();

}
return isr;
}

/*si tiene fechas diferentes te regresa un array 
con cada fecha para posteriormente comparar si solo e sunba fecha te regresa el arreglo con la fecha*/
function date_format(){
  let date_ymd; 
  let array_date;
  if(aux_date.length > 1){

    date_ymd = aux_date.map(item=>{

     for (i in item){

       if( item[i].match(exp_yyyymmdd)){
        fecha = new Date(item[i]);
        month = fecha.getMonth() + 1;
        day = fecha.getDate() + 1;
        year = fecha.getFullYear();


        if(day=='32'){ month = month +1; day = '01';} 

        if(month.toString().length <2 )
          month = '0'+ month;
        if(day.toString().length <2 )
          day = '0'+ day;

      }
      else{
       date_aux = item[i].split(" ")[0].split("-").reverse().join("-");
       fecha = new Date(date_aux);
       month = fecha.getMonth() + 1;
       day = fecha.getDate() + 1;
       year = fecha.getFullYear();

       if(day=='32'){ month = month +1; day = '01';} 

       if(month.toString().length <2 )
        month = '0'+ month;
      if(day.toString().length <2 )
        day = '0'+ day;
    }
    date_ymd = [year,month,day];
  }
  return  date_ymd; 
});
  } else{
    date_ymd = aux_date;
  }
  return  date_ymd; 
}

function data(){

  first_date = date_format();
let aux_firstdate;  //eliminamos el primer elemento del resto
let aux_arrdate;
var result ;

var rfc = aux_rfc[0];
var rn= aux_rn[0];
var numtotal= aux_total[0];

if(first_date.length >1 ){

 aux_firstdate = first_date[0];
 aux_arrdate =  first_date.splice(1);

 var date_inf = aux_arrdate.map(item=>{
  for(i in item){
    aux_firstdate[0] == item[0] ? years = aux_firstdate[0]: years = '*' ;
    aux_firstdate[1] == item[1] ? months = aux_firstdate[1]: months = '*' ;
    aux_firstdate[2] == item[2] ? days = aux_firstdate[2]: days = '*' ;

  }
  return years+'-'+months+'-'+days;

});

 /*eliminamos los elementos repetidos del array*/
 const dataArr = new Set(date_inf);

 result = [...dataArr];

}else{
  aux_firstdate = first_date[0];
  result = aux_firstdate;
}

return result.concat(rfc).concat(rn).concat(numtotal);

}

let respuesta = await alldata();

//HOJA DE TRABAJO
var ws = wb.addWorksheet('consultas');

ws.cell(1,1).string("Fecha-Pago");
ws.cell(1,2).string("Rfc emisor");
ws.cell(1,3).string("Tipo nómina");
ws.cell(1,4).string("Tipo Regimen");
ws.cell(1,5).string("Total");

for (let i = 0; i<respuesta.length;i++){
  let aux = respuesta[i];
  let  aux_pos = aux[i,4];
  let total_90 = aux_pos.toString();
  ws.cell(i+2,1).string(aux[i,0]);
  ws.cell(i+2,2).string(aux[i,1]);
  ws.cell(i+2,3).string(aux[i,2]);
  ws.cell(i+2,4).string(aux[i,3]);
  ws.cell(i+2,5).string(total_90);

  ws.column(i+1).setWidth(20);

}

console.log('Excel generado');

const pathExcel = path.join(__dirname,'excel','ventas.xlsx');

wb.write('consultas.xlsx');

}catch (err) {
  console.error(err);
} finally {
  if (connection) {
    try {
      await connection.close();
    } catch (err) {
      console.error(err);
    }
  }
}

}//termina la funcion run

run();




