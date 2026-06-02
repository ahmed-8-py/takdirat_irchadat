let scrlup = document.getElementById("scrl");
onscroll = function() {
    // إظهار زر التمرير إلى الأعلى عندما يتم التمرير لأسفل الصفحة
    if (window.scrollY > 400) {
        scrlup.style.display = "block";
              

    } else {
        scrlup.style.display = "none";
    }
}


    scrlup.onclick = function() {
    // التمرير إلى أعلى الصفحة
    window.scrollTo({
        top: 0,
        behavior: "smooth" // تمرير ناعم بدلاً من القفز المفاجئ
    });

   
}




//-------------------------------------------------




let container_irchad = document.getElementById("container_irchad");
container_irchad.style.display = "none";
function hide_irchad() {

    let btn_irchad = document.getElementById("hide_irchad");

    if (container_irchad.style.display === "none") {

        container_irchad.style.display = "block";
        btn_irchad.innerHTML = "اخفاء الإرشادات";

    } else {

        container_irchad.style.display = "none";
        btn_irchad.innerHTML = "اظهار الإرشادات";
    }
}


// -------- save notes --------


let datatable_arabic =
JSON.parse(localStorage.getItem("datatable_arabic")) || [];

let datatable_french =
JSON.parse(localStorage.getItem("datatable_french")) || [];

let datatable_english =
JSON.parse(localStorage.getItem("datatable_english")) || [];

let datatable_spanish =
JSON.parse(localStorage.getItem("datatable_spanish")) || [];


show_data();


function get_storage_name() {
    return "datatable_" + document.getElementById("choose_langue").value;
}

    let irchadat = document.querySelectorAll(".irchadat_input");
    let takdirat = document.querySelectorAll(".takdirat_input");

    let min = document.querySelectorAll(".moyenne_min");
    let max = document.querySelectorAll(".moyenne_max");

function save_notes() {

    let storage_name = get_storage_name();

    let datatable = [];




    for (let i = 0; i < irchadat.length; i++) {

        let table = {
            irchadat: irchadat[i].value,
            takdirat: takdirat[i].value,
            min: min[i].value,
            max: max[i].value
        };

        datatable.push(table);
    }

    localStorage.setItem(storage_name, JSON.stringify(datatable));

    
}




function show_data() {

    let storage_name = get_storage_name();

    let data = JSON.parse(localStorage.getItem(storage_name)) || [];

    let irchadat = document.querySelectorAll(".irchadat_input");
    let takdirat = document.querySelectorAll(".takdirat_input");

    let min = document.querySelectorAll(".moyenne_min");
    let max = document.querySelectorAll(".moyenne_max");


    for (let i = 0; i < data.length; i++) {

        irchadat[i].value = data[i].irchadat;
        takdirat[i].value = data[i].takdirat;

        min[i].value = data[i].min;
        max[i].value = data[i].max;
    }
}

// ---------select_langue-----------------


function select_langue() {
    show_data();
}






// ------------------ import exel ------------------






let tbody_excel = document.getElementById("tbody_excel");

let thead_excel  = document.getElementById("thead_excel");


let excelFile = document.getElementById("excelFile");
let uploadIcon = document.getElementById("uploadIcon");

// عند الضغط على الايقونة يفتح اختيار الملف
uploadIcon.onclick = function () {
    excelFile.click();
};


let workbook;

let all_sheets = document.getElementById("all_sheets");

// عند اختيار ملف Excel
excelFile.addEventListener("change", function (e) {

    let file = e.target.files[0];
  // اسم الملف
    let fileName = file.name;

    document.getElementById("file_name").innerHTML = file.name;

    let reader = new FileReader();

    reader.onload = function (event) {

        let data = new Uint8Array(event.target.result);

        workbook = XLSX.read(data, { type: "array" });

        // حذف المحتوى القديم
        all_sheets.innerHTML = "";
      


// إنشاء جدول الأخطاء أولا
let errorDiv = document.createElement("div");
errorDiv.id = "errorDiv";
errorDiv.innerHTML = `
    <h2 >
        جدول الأخطاء
    </h2>

    <table >
        <thead>
            <tr>
                <td>القيمة</td>
                <td>نوع الخطأ</td>
                <td>الفوج التربوي</td>
                <td>التلميذ</td>
            </tr>
        </thead>

        <tbody id="tbody_false"></tbody>
    </table>
`;

all_sheets.appendChild(errorDiv);

// إعادة ربط tbody_false
let tbody_false = document.getElementById("tbody_false");

 let errorCount = 0;



        // المرور على جميع الشيتات
        workbook.SheetNames.forEach(sheetName => {


    // تجاهل الشيت الفارغ الافتراضي
    if (sheetName === "Worksheet") {
        return;
    }
        
        
            // جلب الشيت

            let worksheet = workbook.Sheets[sheetName];

            let jsonData = XLSX.utils.sheet_to_json(worksheet, {
                header: 1 , blankrows: true
            });


            let headers = jsonData[7];
            let kism_shit = jsonData[4];
            // تحويل السطر إلى نص
            let text = kism_shit.join(" ");
            // استخراج ما بعد "الفوج التربوي :"
            let el_fawj = text.split("الفوج التربوي :")[1];

        
            // إنشاء div للشيت
            let sheetDiv = document.createElement("div");
            sheetDiv.id = "sheetdive";
           

            // عنوان الشيت
            let title = document.createElement("h2");
            title.id = "title_sheet";
            title.textContent = el_fawj;
           

            // جدول النتائج
            let table = document.createElement("table");
            table.id = "table_sheet_all";
           

            let thead = document.createElement("thead");
            let tbody = document.createElement("tbody");



            if ( headers.length == 9) {
            thead.innerHTML = `
                <tr>
                    <td>الارشادات</td>
                    <td>التقديرات</td>
                    <td>المعدل</td>
                    <td>الاختبار /20</td>
                    <td>معدل الفروض /20</td>
                    
                    <td>التقييم المستمر /20</td>
                    <td>الاسم</td>
                    <td>اللقب</td>  
                </tr>
            `;
            }
            

            
            if ( headers.length == 10) {
            thead.innerHTML = `
                <tr>
                    <td>الارشادات</td>
                    <td>التقديرات</td>
                    <td>المعدل</td>
                    <td>الاختبار /20</td>
                    <td>معدل الفروض /20</td>
                    <td>أعمال تطبيقية او تعبير شفوي  /20</td>
                    <td>التقييم المستمر /20</td>
                    <td>الاسم</td>
                    <td>اللقب</td>  
                </tr>
            `;
            }
            

              
            if ( headers.length == 7) {
            thead.innerHTML = `
                <tr>
                    
                    <td>تثمين المطالعة او المشاريع</td>
                    
                    <td>الاسم</td>
                    <td>اللقب</td>  
                </tr>
            `;
            }
          

            // المرور على الصفوف
            for (let rowa = 8; rowa < jsonData.length; rowa++) {

                let k = jsonData[rowa];
               

                if ( headers.length == 9) {

                    
            for (let col = 4; col < 7; col++) {

                 let cellValue = jsonData[rowa][col];

                // التحقق من أن القيمة ليست نصية وأنها أكبر من 20
                if (cellValue > 20 || typeof cellValue !== "number" || cellValue == "0" || cellValue == "معفى") {


                 
                        let value = (cellValue === undefined || cellValue === null || cellValue === "")? "قيمة فارغة": cellValue;
                        errorCount ++;
                    tbody_false.innerHTML += `
                        <tr>             
                            <td> ${value}</td>
                            <td> ${headers[col]}</td>
                            <td> ${el_fawj}</td>
                            <td> ${jsonData [rowa] [1] + " " + jsonData [rowa] [2]}</td>
                        </tr>
                    `;


                }
       
       
       

       
            }


                let moyenne = (
                    (Number(k[4]) +
                    Number(k[5]) +
                    (Number(k[6]) * 2)) / 4
                );

                let moyenne_a = isNaN(moyenne)
                    ? "لا يمكن"
                    : moyenne.toFixed(2);

                let result = get_takdir_irchad(moyenne);
                tbody.innerHTML += `
                    <tr>
                        <td>${result.irchadat}</td>
                        <td>${result.takdirat}</td>
                        <td>${moyenne_a}</td>
                        <td>${k[6] ?? ""}</td>
                        <td>${k[5] ?? ""}</td>
                        
                        <td>${k[4] ?? ""}</td>
                        <td>${k[2] ?? ""}</td>
                        <td>${k[1] ?? ""}</td>            
                    </tr>
                `;
            }


            // +++++++++++++++++++++++++++++++++++++++++++++

        

                if ( headers.length == 10) {

                    
            for (let col = 4; col < 8; col++) {

                 let cellValue = jsonData[rowa][col];

                // التحقق من أن القيمة ليست نصية وأنها أكبر من 20
                if (cellValue > 20 || typeof cellValue !== "number" || cellValue == "0" || cellValue == "معفى") {


                 
                        let value = (cellValue === undefined || cellValue === null || cellValue === "")? "قيمة فارغة": cellValue;
                        errorCount ++;
                    tbody_false.innerHTML += `
                        <tr>             
                            <td> ${value}</td>
                            <td> ${headers[col]}</td>
                            <td> ${el_fawj}</td>
                            <td> ${jsonData [rowa] [1] + " " + jsonData [rowa] [2]}</td>
                        </tr>
                    `;


                }
       
       
       

       
            }


                let moyenne = (
                    (Number(k[4]) +
                    Number(k[5]) +
                    Number(k[6]) +
                    (Number(k[7]) * 2)) / 5
                );

                let moyenne_a = isNaN(moyenne)
                    ? "لا يمكن"
                    : moyenne.toFixed(2);

                let result = get_takdir_irchad(moyenne);
                tbody.innerHTML += `
                    <tr>
                        <td>${result.irchadat}</td>
                        <td>${result.takdirat}</td>
                        <td>${moyenne_a}</td>
                        <td>${k[7] ?? ""}</td>
                        <td>${k[6] ?? ""}</td>
                        <td>${k[5] ?? ""}</td>
                        <td>${k[4] ?? ""}</td>
                        <td>${k[2] ?? ""}</td>
                        <td>${k[1] ?? ""}</td>            
                    </tr>
                `;
            }



            // +++++++++++++++++++++++++++++++++++++++++++++



                if ( headers.length == 7) {

                    
            for (let col = 4; col < 5; col++) {

                 let cellValue = jsonData[rowa][col];

                // التحقق من أن القيمة ليست نصية وأنها أكبر من 20
                if (cellValue > 20 || typeof cellValue !== "number" || cellValue == "0" || cellValue == "معفى") {


                 
                        let value = (cellValue === undefined || cellValue === null || cellValue === "")? "قيمة فارغة": cellValue;
                        errorCount ++;
                    tbody_false.innerHTML += `
                        <tr>             
                            <td> ${value}</td>
                            <td> ${headers[col]}</td>
                            <td> ${el_fawj}</td>
                            <td> ${jsonData [rowa] [1] + " " + jsonData [rowa] [2]}</td>
                        </tr>
                    `;


                }
       
       
       

       
            }


                tbody.innerHTML += `
                    <tr>
                        
                        <td>${k[4] ?? ""}</td>
                        <td>${k[2] ?? ""}</td>
                        <td>${k[1] ?? ""}</td>            
                    </tr>
                `;
            }




            // +++++++++++++++++++++++++++++++++++++++++++++
}


       
            table.appendChild(thead);
            table.appendChild(tbody);

            sheetDiv.appendChild(title);
            sheetDiv.appendChild(table);

            all_sheets.appendChild(sheetDiv);

            
        });
       
if (errorCount ===0) {
                        tbody_false.innerHTML += `
                        <tr>             
                            <td>/</td>
                            <td>/</td>
                            <td>/</td>
                            <td>/</td>
                        </tr>
                    `;
    
}

    };

     reader.readAsArrayBuffer(file);


     

});



function get_takdir_irchad(moyenne) {

    let storage_name = get_storage_name();

    let data = JSON.parse(localStorage.getItem(storage_name)) || [];

    moyenne = Number(moyenne);

    for (let i = 0; i < data.length; i++) {

        let min = Number(data[i].min);
        let max = Number(data[i].max);

        // تصحيح تلقائي إذا كانت القيم معكوسة
        if (min > max) {
            [min, max] = [max, min];
        }

        if (moyenne >= min && moyenne <= max) {

            return {
                takdirat: data[i].takdirat,
                irchadat: data[i].irchadat
            };
        }
    }

    return {
        takdirat: "/",
        irchadat: "/"
    };
}















async function export_excel() {

    let file = excelFile.files[0];
    let fileName = file.name;




    if (!file) {
        alert("اختر ملف Excel");
        return;
    }
    
    let arrayBuffer = await file.arrayBuffer();

    // فتح الملف الأصلي نفسه
    let workbook = new ExcelJS.Workbook();

    await workbook.xlsx.load(arrayBuffer);

    // المرور على كل الشيتات
    workbook.eachSheet((worksheet, sheetId) => {

        // تجاهل الشيت الفارغ
        if (worksheet.name === "Worksheet") return;


    let headersCount = worksheet.getRow(8).cellCount;



        // بداية التلاميذ
        for (let row = 9; row <= worksheet.rowCount; row++) {
//+++++++++++++++++++++++

        if ( headersCount == 9) {

            // القيم الأصلية
            let cc = worksheet.getCell("E" + row).value;
            let dd = worksheet.getCell("F" + row).value;
            let ee = worksheet.getCell("G" + row).value;

            let moyenne = (
                Number(cc) +
                Number(dd) +
                (Number(ee) * 2)
            ) / 4;

            if (isNaN(moyenne)) continue;

            moyenne = Number(moyenne.toFixed(2));

            let result = get_takdir_irchad(moyenne);

            // ======================
            // الكتابة داخل الخلايا الأصلية
            // ======================

           // worksheet.getCell("I" + row).value = moyenne;

            worksheet.getCell("H" + row).value =
                result.takdirat;

            worksheet.getCell("I" + row).value =
                result.irchadat;
        }

//====================

        
        if ( headersCount == 10) {

            // القيم الأصلية
            let cc = worksheet.getCell("E" + row).value;
            let dd = worksheet.getCell("F" + row).value;
            let ee = worksheet.getCell("G" + row).value;
            let ff = worksheet.getCell("H" + row).value;

            let moyenne = (
                Number(cc) +
                Number(dd) +
                Number(ee) +
                (Number(ff) * 2)
            ) / 5;

            if (isNaN(moyenne)) continue;

            moyenne = Number(moyenne.toFixed(2));

            let result = get_takdir_irchad(moyenne);

            // ======================
            // الكتابة داخل الخلايا الأصلية
            // ======================

           

            worksheet.getCell("I" + row).value =
                result.takdirat;

            worksheet.getCell("J" + row).value =
                result.irchadat;
        }

//+++++++++++++++++++++++





        }



    });

    // إنشاء الملف النهائي
    let buffer = await workbook.xlsx.writeBuffer();
    

    saveAs(
        new Blob([buffer]),
        "az" + `${fileName}`
    );
}







