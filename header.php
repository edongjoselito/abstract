<!DOCTYPE html>
<html lang=en>
  <head>
    <title>CRUD</title>
    <link rel="stylesheet" href="style.css" />
    <script src="jquery.js" type="text/javascript"></script>
    <script>
      $(document).ready(function(){
        $("#city option").hide();
        $("#brgy option").hide();

        $("#province").change(function(){
           var val=$(this).val();
           $("#city option").hide();
           $("#city").val("");
           $("#city [data-city='" + val + "']").show();
           $("#city").change();
          });
          
          $("#city").change(function(){
           var val=$(this).find(":selected").prop("id");
           $("#brgy option").hide();
           $("#brgy").val("");
           $("#brgy [data-brgy='" + val + "']").show();
          });

        
        });
    </script>
  </head>
  <body>

    <div class="wrap">
      <div class="left">