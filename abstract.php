<?php include("connection.php"); ?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link rel="stylesheet" href="style.css">
    <link href="https://db.onlinewebfonts.com/a/0nH393RJctHgt1f2YvZvyruY" rel="stylesheet" type="text/css"/>
    <title>ABSTRACT OF QUOTATION</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.0/xlsx.full.min.js"></script>
</head>
<body>


<div class="page-number"></div>

<?php $ren_id=$_GET['id']; ?>

<div class="renren">
    <form action="abstract.php?id=<?= $ren_id; ?>" method="POST" class="styled-form">
        <select class="styled-select" name="lot">
            <?php 
                $sql = "SELECT * FROM proc_pr WHERE actID=$ren_id group by lot order by lot";
                $row_set = mysqli_query($con, $sql);
                while($rows = mysqli_fetch_assoc($row_set)){
            ?>
            <option value="<?= $rows['lot']; ?>"><?= $rows['lot']; ?></option>

            <?php } ?>
        </select>
        <input type="submit" value="Submit" name="Submit_lot" class="styled-submit">
    </form>
</div>

<?php if(isset($_POST['Submit_lot'])){$cl=$_POST['lot'];?>


<div class="hwrap">
    <img class="logo" src="ke.png" alt="">
        <p class="textwrap">
        <span class="rp">Republic of the Philippines</span>
            <span class="de">Department of Education</span>
            <span class="r">Region XI</span>
            <span class="r">School Division of Davao Oriental</span>
            <span class="r">Mati City</span>
        </p>
        <h1 class="na">ABSTRACT OF QUOTATION</h1>
</div>


<div class="imgwrap">
<img class="h_image" src="heading.jpg" alt="">
</div>




<?php 
    $sql = "SELECT * FROM proc_quotation WHERE actID=$ren_id and lot='".$cl."' group by supplier";
    $row_set = mysqli_query($con, $sql);
    $cr = mysqli_num_rows($row_set);


    $sql = "SELECT * FROM ls_activities where actID=$ren_id";
    $row_set = mysqli_query($con, $sql);
    $act = mysqli_fetch_assoc($row_set) 
?>

<table class="guaposirenren">
    <tr>
            <td class="topcont ic">Mode of Procurement</td>
            <td class="text-center bgblue">NP-Small Value Procurement</td>
        </tr>
        <tr>
            <td class="topcont">Quotation No.:</td>
            <td class="text-center bgblue">24-08-0008</td>
        </tr>
        <tr>
            <td class="topcont">Date of Opening</td>
            <td class="text-center">9/3/2024</td>
        </tr>
</table>

<table class="ab_report">
    <tr>
        <th rowspan="2">ITEM NO.</th>
        <th rowspan="2">UNIT OF ISSUE</th>
        <th rowspan="2">ITEM & DESCRIPTION</th>
        <th rowspan="2">QTY</th>
        <th rowspan="2">Estimate<br /> Cost</th>
        <?php
              $sql = "SELECT * FROM proc_quotation WHERE actID=$ren_id and lot='".$cl."' group by supplier";
              $row_set = mysqli_query($con, $sql);
              while($row = mysqli_fetch_assoc($row_set)){ ?>

            <th colspan="2"><?= $row['supplier']; ?></th>

        <?php } ?>
    </tr>
    <tr>
    <?php
              $sql = "SELECT * FROM proc_quotation WHERE actID=$ren_id and lot='".$cl."' group by supplier";
              $row_set = mysqli_query($con, $sql);
              while($row = mysqli_fetch_assoc($row_set)){ ?>
        <th>Unit Price</th>
        <th>Total Price</th>
    <?php } ?>
    </tr>
    <?php
              $c=1;
              $cid=1;
              $ec=0;
              $sql = "SELECT * FROM proc_pr WHERE actID=$ren_id and lot='".$cl."'";
              $row_set = mysqli_query($con, $sql);
              while($rows = mysqli_fetch_assoc($row_set)){ 
                $pr_id = $rows['prID'];
                $sub_ec = $rows['est_cost']*$rows['qty'];
                $fec = $ec+=$sub_ec;
    ?>
    <tr>
        <td class="text-center"><?= $c++; ?></td>
        <td><?= $rows['unit']; ?></td>
        <td><?= $rows['item_description']; ?></td>
        
        <td class="text-center"><?= $rows['qty']; ?></td>
        <td class="text-center"><?= number_format($rows['est_cost']); ?></td>
        <?php
            $sql = "SELECT * FROM proc_quotation WHERE actID=$ren_id and lot='".$cl."' GROUP BY supplier";
            $result_set = mysqli_query($con, $sql);

            // Initialize a variable to track the lowest q_amount
            $min_q_amount = PHP_INT_MAX; // Start with the largest possible integer

            // First pass to find the minimum q_amount
            while($row = mysqli_fetch_assoc($result_set)) {
                $prs = str_replace("'", "\'", $row['supplier']);
                
                $sql = "SELECT * FROM proc_quotation WHERE prID=$pr_id and lot='".$cl."' and supplier='".$prs."'";
                $pro_set = mysqli_query($con, $sql);
                $pr = mysqli_fetch_assoc($pro_set);

                if ($pr['q_amount'] > 0.01 && $pr['q_amount'] < $min_q_amount) {
                    $min_q_amount = $pr['q_amount']; // Update min_q_amount
                }
            }

            // Reset result_set for the second pass
            mysqli_data_seek($result_set, 0); // Reset the result set pointer

            // Second pass to display the values and highlight the minimum
            while($row = mysqli_fetch_assoc($result_set)) {
                $prs = str_replace("'", "\'", $row['supplier']);
                
                $sql = "SELECT * FROM proc_quotation WHERE prID=$pr_id and lot='".$cl."' and supplier='".$prs."'";
                $pro_set = mysqli_query($con, $sql);
                $pr = mysqli_fetch_assoc($pro_set);
                
                // Determine if this is the minimum q_amount for highlighting
                $is_lowest = $pr['q_amount'] == $min_q_amount;
            ?>
                <td class="text-center" style="<?php echo $is_lowest ? 'background-color: #718093; color: #fff' : ''; ?>">
                    <?php if($pr['q_amount'] == 0.01) {
                        echo "N/A";
                    } else {
                        echo number_format($pr['q_amount']);
                    } ?>
                </td>
                <td class="text-center" style="<?php echo $is_lowest ? 'background-color: #718093; color: #fff' : ''; ?>">
                    <?php if($pr['q_amount'] == 0.01) {
                        echo "N/A";
                    } else {
                        echo number_format($pr['q_amount'] * $rows['qty'], 2);
                    } ?>
                </td>
            <?php 
            } 
            ?>


    </tr>
    
    
    <?php } ?>

    

    <tr>
        <td>&nbsp;</td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <?php
              $sql = "SELECT * FROM proc_quotation WHERE actID=$ren_id and lot='".$cl."' group by supplier";
              $result_set = mysqli_query($con, $sql);
              while($row = mysqli_fetch_assoc($result_set)){ 
                
        ?>
            <td class="text-center"></td>
            <td class="text-center"></td>
      
        <?php } ?>
    </tr>
    <tr>
        <td colspan="5" style="padding:8px"> TOTAL AMOUNT WITH LOWEST PRICE QUOTATION: </td>
        
        <?php
              
              $sql = "SELECT * FROM proc_quotation WHERE actID=$ren_id and lot='".$cl."' group by supplier";
              $result_set = mysqli_query($con, $sql);
              while($row = mysqli_fetch_assoc($result_set)){ 
                $q = str_replace("'", "\'", $row['supplier']);
                
        ?>

        <?php $sql = "SELECT * FROM proc_quotation WHERE actID=$ren_id  and supplier='".$q."'";
              $row_set = mysqli_query($con, $sql);
              $ren=0;
              $renren_guapo=0;
              $has_zero_point_one = false; // Flag to track if 0.01 is found
              while($pra = mysqli_fetch_assoc($row_set)){ 
                $id = $pra['prID'];

                $sql = "SELECT * FROM proc_pr WHERE prID=$id";
                $pro_set = mysqli_query($con, $sql);  
                $qty = mysqli_fetch_assoc($pro_set);

                $sql = "SELECT * FROM proc_quotation WHERE actID=$ren_id and supplier='".$q."' LIMIT 1" ;
                $sup_set = mysqli_query($con, $sql);
                $sup = mysqli_fetch_assoc($sup_set);
                

            ?>
            
            <?php 
                $ta= $pra['q_amount']*$qty['qty']; 

                if ($pra['q_amount'] == 0.01) {
                    $has_zero_point_one = true; // Set the flag if 0.01 is found
                }
            ?>

            
            <?php $ren+=$ta; $renren_guapo+=$ta; ?>


            <?php }  ?>    
        
            <td class="text-center" colspan="2" <?php  if($fec <= $ren){echo 'style="color:red"';} ?>>
            <?php $sup_name = $sup['supplier']; ?>

            
            <?php 
            if (!$has_zero_point_one) {
                $names[$sup_name] = $ren;
            }

            //$ivyclaire[$sup_name]=$ren;

            //echo $ivyclaire;
            ?>




            <b><?= number_format($renren_guapo, 2); ?></b>

            </td>
            

            
            
        <?php } ?>
        
        <?php if(!empty($names)){
            $minValue = min($names);

            // Find the key associated with the minimum value
            $minKey = array_search($minValue, $names);
            }
        ?>



 
      
    </tr>
    <tr>
        <td colspan="5" class="tb"> REMARKS (RESPONSIVE, PASSED, OR NO-RESPONSIVE) </td>
        <?php
              $sql = "SELECT * FROM proc_quotation WHERE actID=$ren_id and lot='".$cl."' group by supplier";
              $result_set = mysqli_query($con, $sql);
              while($row = mysqli_fetch_assoc($result_set)){ 
                
        ?>
            <td class="text-center" colspan="2"></td>
      
        <?php } ?>
    </tr>
    <tr>
        <td colspan="5" class="tb"> OTHER REMARKS</td>
        <?php
              $sql = "SELECT * FROM proc_quotation WHERE actID=$ren_id and lot='".$cl."' group by supplier";
              $result_set = mysqli_query($con, $sql);
              $row_count = mysqli_num_rows($result_set);
              while($row = mysqli_fetch_assoc($result_set)){ 
                
        ?>
            <td class="text-center"></td>
            <td class="text-center"></td>
      
        <?php } ?>
    </tr>
    <tr>
        <td colspan="2" class="tb">Purpose</td>
        <td colspan="<?= $row_count*2+3; ?>"><strong><?= $act['act_title']; ?></strong></td>
    </tr>
    <tr>
        <td colspan="3" class="tb">Approved  Budget for the Contract:</td>
        <td colspan="<?= $row_count*2+2; ?>">  <b style="margin-left:2%;"><?= number_format($fec, 2); ?></b></td>
    </tr>
    <tr>
        <td colspan="<?= $row_count*2+5; ?>" style="border:0; background-color:#3b76fd; font-weight:bold; padding:10px; color:#fff">Note: The highlighted amount in "Unit Price" comumn is the lowest qouted price for the item. Absence of a highlight means the item for procurement is 'FAILED.'</td>
    </tr>
    <tr>
        <td colspan="<?= $row_count*2+5; ?>" style="border:0; padding:10px; text-indent: 50px;">
            The undersigned members of the committee on awards hereby certify that we certify that we have this 3rd day of September, 2024 opened bids of the above items that the prices by each dealer are as appearing in their quatation and that the lowest and most advantageous price for the government in accordance with samples or specification submitted and recommend to be awarded to:
        </td>
    </tr>
    <tr>
        <td colspan="<?= $row_count*2+5; ?>" style="text-align:center; background-color:yellow; padding:8px 0; font-weight:bold">

        <?php if(!empty($minKey)){echo $minKey;} ?>
        </td>
    </tr>




</table>

<div class="sign">
    <p>EMMA O. RABUYA <span>BAC Member</span></p>
    <p>NANCY P. SUMAGAYSAY, PHD<span>BAC Member</span></p>
    <p>YVETTE M. CELMAR, PHD <span>BAC Member</span></p>
    <p>BERNABE M. BASILISCO<span>BAC Member</span></p>
    <p>ERNESTO H. CABANES<span>Acting Chairperson</span></p>
</div>


<?php } ?>

</body>
</html>
