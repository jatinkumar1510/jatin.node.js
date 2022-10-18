// const NAME="Jatin";
// let num=0;
// const studentObject=
// {
//         NAME:"Jatin",
//         branch :"cse",
//         CGPA :10,
//         "fav hobby":"cricket",
//         address:
//         {
//             city:"namchi",
//             country: "India",
//             region: "Asia", 
//         }
// };
// studentObject.CGPA=9;
// const {NAME}=studentObject;
// console.log(NAME)
// console.log(studentObject.CGPA)

const parser = require('simple-excel-to-json')
var json2xls = require('json2xls');
const doc = parser.parseXls2Json('./Assignment.xlsx')[0]; 
const fs = require("fs");

console.log(doc)

const sortByColumn = CGPA => {
    if (Assignment) {
      CGPA--;
      const sortFunction = (a, b) => {
        if (a[CGPA] === b[CGPA]) {
          return 0;
        }
        else {
          return (a[CGPA] < b[CGPA]) ? -1 : 1;
        }
      }
      let rows = [];
      for (let i = 1; i <= Assignment.actualRowCount; i++) {
        let row = [];
        for (let j = 1; j <= Assignment.columnCount; j++) {
          row.push(Assignment.getRow(i).getCell(j).value);
        }
        rows.push(row); 
      }
      rows.sort(sortFunction);
      // Remove all rows from Assignment then add all back in sorted order
      Assignment.spliceRows(1, Assignment.actualRowCount);
      // Note Assignment.addRows() may add them to the end of empty rows so loop through and add to beginnning
      for (let i = rows.length; i >= 0; i--) {
        Assignment.spliceRows(1, 0, rows[i]);
      }
    }
  }


const xlsData = json2xls()

  fs.writeFileSync('sorted.xlsx', xlsData,'binary');





























// console.log(doc)
// () => {}
// const avg = doc.reduce((prevalue,current)=>{
//     console.log(prevalue)
//     return prevalue
// }, 0)
// const totalCGPA = doc.reduce((prevalue,current)=>
// {
//     prevalue+=current.CGPA
//     return prevalue
// }, 0);

// const avg = totalCGPA / doc.length;
// console.log(totalCGPA)
// const docWithAverage = [...doc];
// const CGPAdoc=docWithAverage.map((student)=>
// {
//     if(student.CGPA>9.5)
//     {
//         student.CGPA="A+"
//     }
//     else if(student.CGPA > 9.2 && student.CGPA < 9.5)
//     {
//         student.CGPA="A"
//     }
//     else if(student.CGPA > 8.8 && student.CGPA < 9.2)
//     {
//         student.CGPA="B"
//     }
//     else if(student.CGPA > 8.6 && student.CGPA < 8.8)
//     {
//         student.CGPA="C"
//     }
//     else if(student.CGPA>7 && student.CGPA<8.6)
//     {
//         student.CGPA="D"
//     }
//     else if(student.CGPA>6 && student.CGPA<7)
//     {
//         student.CGPA="E"
//     }
//     else 
//     {
//         student.CGPA="fail"
//     }
    
//     return student
// })

// const fileredDocument=CGPAdDocuments.filter((student))
// if(student.CGPA>8)
// {
//     return true
// }
// else
// {
//     return false
// }


// docWithAverage.push({CGPA:avg, NAME:"avg"})
// const xlsData = json2xls(docWithAverage)
// console.log(avg)
// fs.writeFileSync('data.xlsx', xlsData,'binary');
// fs.writeFileSync('filtered.xlsx', xlsData,'binary');