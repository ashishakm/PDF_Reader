<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ReadPDF.aspx.cs" Inherits="PDF_Demo.View.ReadPDF" %> 

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">    
    <title></title>   
  <style type="text/css">
        #dropOnMe {
            width: 210px;
            height: 136px;
            padding: 10px;

            border: 2px dashed gray;
            background-color: lightgray;
        }
    </style>
   
    <script src="../Scripts/jquery-1.9.1.js"></script>
   
</head>
<body>
    <form id="form1" runat="server">
        <div style="margin-left:300px;"> 
            <asp:FileUpload ID="FileUpload1" runat="server" /> <br /> <br />          
            <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="Click Here upload pdf file" />
        </div>     
    </form>
       <h3>Drop Files on Me</h3>
    <div id="dropOnMe" draggable="false"></div>
    <div id="fileCount" draggable="false"></div>
    <input id="upload" draggable="false" type="button"
        value="Upload Selected Files" />
    <div draggable="false">
        <ol draggable="false" id="myFileList"></ol>
    </div>
     <script>
         $(document).ready(function () {
             // this function runs when the page loads to set up the drop area

             // Check if window.fileReader exists to make sure the browser 
             // supports file uploads
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
                 if (!filelist) return;  // if null, exit now

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
         });
    </script>
</body>
    
</html>
