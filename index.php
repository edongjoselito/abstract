<?php include("connection.php"); ?>
<?php $page="home"?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>ABSTRACT OF QUOTATION</title>
    <link rel="stylesheet" href="style.css">
    <link rel="stylesheet" href="https://cdn.datatables.net/1.11.5/css/jquery.dataTables.min.css">
    <link rel="stylesheet" href="https://cdn.datatables.net/1.11.5/css/dataTables.bootstrap4.min.css">
</head>
</head>
<body>

    <div class="lc">
      <div class="nav">
        <a href="#">HOME</a>
        <a <?php if($page == "home"){echo "class='active'";} ?>href="#">ITEMS</a>
      </div>
       
    </div>

    <div class="rc">
      <div class="table_rap">
        <table id="example" class="display my-custom-table" style="width:100%">
            <thead>
                <tr>
                    <th>Name</th>
                    <th>Action</th>
                </tr>
            </thead>
            <tbody>
            <?php
              $sql = "SELECT * FROM ls_activities";
              $row_set = mysqli_query($con, $sql);
              while($row = mysqli_fetch_assoc($row_set)){ ?>

              <tr>
                <td><?= $row['act_title']; ?></td>
                <td>
                  <a target="_blank" href="abstract.php?id=<?= $row['actID']; ?>">View</a>
                </td>
              </tr>
              <?php } ?>
          </tbody>
      </table>
    </div>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.datatables.net/1.11.5/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.datatables.net/1.11.5/js/dataTables.bootstrap4.min.js"></script>
    <script>
        $(document).ready(function() {
            $('#example').DataTable();
        });
    </script>
    </div>

</body>
</html>