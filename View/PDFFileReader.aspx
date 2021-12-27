    <%@ Page Language="C#" AutoEventWireup="true" CodeBehind="PDFFileReader.aspx.cs" Inherits="PDF_Demo.View.PDFFileReader" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
        <style>
            * {box-sizing: border-box}

            /* Set height of body and the document to 100% */
            body, html {
              height: 100%;
              margin: 0;
              font-family: Arial;
            }

            /* Style tab links */
            .tablink {
              background-color: #555;
              color: white;
              float: left;
              border: none;
              outline: none;
              cursor: pointer;
              padding: 14px 16px;
              font-size: 17px;
              width: 25%;
            }

            .tablink:hover {
              background-color: #777;
            }

            /* Style the tab content (and add height:100% for full page content) */
            .tabcontent {
              color: white;
              display: none;
              padding: 100px 20px;
              height: 100%;
            }

            #PDFType1 {background-color: lightblue;}
            #PDFType2 {background-color: lightgreen;}
           
       </style>

    <style type="text/css">
        #dropOnMe {
            width: 550px;
            height: 250px;
            padding: 10px;
            border: 2px dashed gray;
            background-color: lightgray;
            color:black;
        }
        #dropOnMe1 {
            width: 550px;
            height: 250px;
            padding: 10px;
            border: 2px dashed gray;
            background-color: lightgray;
             color:black;
        }
    </style>
   
    <script src="../Scripts/jquery-1.9.1.js"></script>
</head>
<body>
    
        <%-----Tabular part-----%>
        <button class="tablink" onclick="openPage('PDFType1', this, 'blue')">PDF Type1</button>
        <button class="tablink" onclick="openPage('PDFType2', this, 'green')" >PDF Type2</button>
        
        <%-----Body part-----%>
    <div>
        <div id="PDFType1" class="tabcontent">
          <h3>PDF Type 1</h3>
           <div>
            <h3>Drop Files on Box</h3>
            <div id="dropOnMe" draggable="false"></div>
            <div id="fileCount" draggable="false"></div>
            <input id="upload" draggable="false" type="button"value="Upload Selected Files" />
            <div draggable="false">
            <ol draggable="false" id="myFileList"></ol>
        </div>
        </div>
        </div>

        <div id="PDFType2" class="tabcontent">
          <h3>PDF Type 2</h3>
           <div>
            <h3>Drop Files on Box</h3>
            <div id="dropOnMe1" draggable="false"></div>
            <div id="fileCount1" draggable="false"></div>
            <input id="upload1" draggable="false" type="button"value="Upload Selected Files" />
            <div draggable="false">
            <ol draggable="false" id="myFileList1"></ol>
        </div>
        </div>
       </div>
        <%------------------------%>
       
     <script>
         function openPage(pageName, elmnt, color) {
             var i, tabcontent, tablinks;
             tabcontent = document.getElementsByClassName("tabcontent");
             for (i = 0; i < tabcontent.length; i++) {
                 tabcontent[i].style.display = "none";
             }
             tablinks = document.getElementsByClassName("tablink");
             for (i = 0; i < tablinks.length; i++) {
                 tablinks[i].style.backgroundColor = "";
             }
             document.getElementById(pageName).style.display = "block";
             elmnt.style.backgroundColor = color;
         }
 
         /* Code for file upload */
         $(document).ready(function () {    
             if (typeof (window.FileReader) == 'undefined') {
                 alert('Browser does not support HTML5 file uploads!');
             }

             dropOnMe.addEventListener("drop", dropHandler, false);

             dropOnMe.addEventListener("dragover", function (ev) {
                 $("#dropOnMe").css("background-color", "lightgoldenrodyellow;");
                 ev.preventDefault();
             }, false);

             function dropHandler(ev) {
                 // Prevent default processing.
                 ev.preventDefault();

                 // Get the file(s) that are dropped.
                 var filelist = ev.dataTransfer.files;
                 if (!filelist) return;  // if null, do not do anything.

                 $("#dropOnMe").text(filelist.length +
                     " file(s) selected for uploading!");

                 $("#upload").click(function () {
                     var data = new FormData();
                     for (var i = 0; i < filelist.length; i++) {
                         data.append(filelist[i].name, filelist[i]);                          
                     }
                    
                     $.ajax({
                         type: "POST",
                         url: "FileUpload.ashx",
                         contentType: false,
                         processData: false,
                         data: data,
                         success: function (result) {
                             alert(result);
                         },
                         error: function () {
                             alert("There was error uploading files!"); 
                         }
                     });
                 });

             }

             dropOnMe.addEventListener("dragend", function (ev) {
                 $("#dropOnMe").css("background-color", "lightgray;");
                 $("#dropOnMe").text("");
                 $("upload").click(function () { });
                 ev.preventDefault();
             }, false);


         /* Code for file upload 1 */
             dropOnMe1.addEventListener("drop", dropHandler1, false);

             dropOnMe1.addEventListener("dragover", function (ev) {
                 $("#dropOnMe1").css("background-color", "lightgoldenrodyellow;");
                 ev.preventDefault();
             }, false);

             function dropHandler1(ev) {
                 // Prevent default processing.
                 ev.preventDefault();

                 // Get the file(s) that are dropped.
                 var filelist = ev.dataTransfer.files;
                 if (!filelist) return;  // if null, do not do anything.

                 $("#dropOnMe1").text(filelist.length +
                     " file(s) selected for uploading!");

                 $("#upload1").click(function () {
                     var data = new FormData();
                     for (var i = 0; i < filelist.length; i++) {
                         data.append(filelist[i].name, filelist[i]);
                     }

                     $.ajax({
                         type: "POST",
                         url: "FileUpload1.ashx",
                         contentType: false,
                         processData: false,
                         data: data,
                         success: function (result) {
                             alert(result);
                         },
                         error: function () {
                             alert("There was error uploading files!");
                         }
                     });
                 });

             }

             dropOnMe1.addEventListener("dragend", function (ev) {
                 $("#dropOnMe1").css("background-color", "lightgray;");
                 $("#dropOnMe1").text("");
                 $("upload1").click(function () { });
                 ev.preventDefault();
             }, false);
         });

        
        
    </script>
  
        </div>
    
</body>
</html>
