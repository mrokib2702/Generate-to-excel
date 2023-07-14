<?php  

    #===================================================|
    # Please DO NOT modify this information :               
    #---------------------------------------------------|
    # @Author       : RKB
    # @Project      : CONSOLE FILE EXCEL 
    # 
    # Sesuaikan user, password dan nama dabatabse 
    # pada masing-masing server 
    #===================================================|


    ini_set('memory_limit', '-1');

    shell_exec('start cmd.exe @cmd /k "D:\CONSOLE_DATA\CONNECT_ANYCONNECT.bat" exit');

    $server = "<YOUR_SERVER_NAME>"; 
    $conn = sqlsrv_connect( $server, array(  "Database"=>"<YOUR_LOCAL_DB_NAME>", "UID"=>"<YOUR_USER_DB>", "PWD"=>"YOUR_PASSWORD_DB" ) );

    if( !$conn ) {
        echo "Connection Failed \n";
    }  
  
    $stmt = sqlsrv_query( $conn, "SELECT 
            TOP 1
            convert(varchar, TGL_TARIK, 20) TGL_TARIK,
            TGL_TARIK TGL_TARIK_REAL
            FROM
            [XX.XX.XX.XX].[<NAMA_DB>].[dbo].[<NAMA_TABLE>]" , array(), array( "Scrollable" => SQLSRV_CURSOR_KEYSET ));  
    
    while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_ASSOC) ) {
        $TGL_TARIK =  $row['TGL_TARIK'];
        $TGL_TARIK_REAL =  $row['TGL_TARIK_REAL'];
    }

    echo "TANGGAL TARIK : ".$TGL_TARIK."\n\n";

    $stmt2 = sqlsrv_query( $conn, "SELECT 
            TOP 1
            TGL_TARIK CAB_TGL_TARIK,
            REGION
            FROM
            [LOG_JOB] 
            WHERE  convert(varchar, TGL_TARIK, 20) = '$TGL_TARIK' AND REGION = 'ALL' " , array(), array( "Scrollable" => SQLSRV_CURSOR_KEYSET ));  


    $row_count = sqlsrv_num_rows( $stmt2 );  
  
    if ($row_count === false){
      echo "\nerror\n";  
    }
    else if ($row_count < 1) {
        //echo "\n$row_count\n";


        require_once "xlsxwriter.class.php";
        $writer = new XLSXWriter();
        $header_style = array(
            'font'=>'Calibri',
            'font-size'=>11, 
            'wrap_text'=>false, 
            'border'=>'left,right,top,bottom',
            'border-style'=>'medium', 
            'border-color'=>'#0000FF', 
            'valign'=>'top', 
            'color'=>'#FFFFFF', 
            'fill'=>'#0000FF',
            'auto_filter'=>true,
        );
    
        $sheet_options = array(
            'autofilter'=>true,
            'freeze_pane' =>array(1,10)
        );
    
        $header = array(
            'Tgl_Tarik'=>'string',
            //'WilayahID'=>'string',
            'WilayahName'=>'string',
            'RegionID'=>'string',
            //'OurBranchIDPNM'=>'string',
            'Initial'=>'string',
            //'BranchName'=>'string',
            'RegionName'=>'string',
            //'AreaID'=>'string',
            'AreaName'=>'string',
            'UnitID'=>'string',
            'UnitName'=>'string',
            'UnitInitial'=>'string',
            //'Kota'=>'string',
            'GroupID'=>'string',
            'GroupName'=>'string',
            'ClientID'=>'string',
            'Name'=>'text',
            'AccountID'=>'string',
            'LoanSeries'=>'string',
            'ProductID'=>'string',
            'InstallmentAmount'=>'#,##0',
            'FirstDisbursementDate'=>'string',
            'OutStandingPrincipal'=>'#,##0',
            'OutStandingInterest'=>'#,##0',
            'ODPrincipal'=>'#,##0',
            'ODInterest'=>'#,##0',
            'NoOfArrearDays'=>'#,##0',
            'DisbursedAmount'=>'#,##0',
            'Term'=>'#,##0',
            'MeetingDayID'=>'#,##0',
            'LoanTypeID'=>'string',
            'RepaymentFrequencyID'=>'string',
            'RepaymentTerm'=>'#,##0',
            'LastPaidDate'=>'string',
            'LastInstallmentNo'=>'#,##0',
            'PrincipalPaid'=>'#,##0',
            'InterestPaid'=>'#,##0',
            'UnearnPrincipal'=>'#,##0',
            'UnearnInterest'=>'#,##0',
            'PrincipalDue'=>'#,##0',
            'InterestDue'=>'#,##0',
            'InstallmentNo'=>'#,##0',
            'InstallmentDueDate'=>'string',
            'PaymentNo'=>'#,##0',
            'InterestAmount'=>'#,##0',
          

        );
    
        $SVR = "<NAMA SERVER>"; 
        $connectionInfo = array( "Database"=>"<NAMA_DB>", "UID"=>"<USER_DB>", "PWD"=>"<PASSWORD_DB>");

        $conn_SVR = sqlsrv_connect( $SVR, $connectionInfo);
        if( !$conn_SVR ) {
            echo "KONEKSI KE DW GAGAL BRO \n\n";
            die( print_r( sqlsrv_errors(), true));
        }
    

        function cleanString($text) {
            $utf8 = array(
                '/[áàâãªä]/u'   =>   'a',
                '/[ÁÀÂÃÄ]/u'    =>   'A',
                '/[ÍÌÎÏ]/u'     =>   'I',
                '/[íìîï]/u'     =>   'i',
                '/[éèêë]/u'     =>   'e',
                '/[ÉÈÊË]/u'     =>   'E',
                '/[óòôõºö]/u'   =>   'o',
                '/[ÓÒÔÕÖ]/u'    =>   'O',
                '/[úùûü]/u'     =>   'u',
                '/[ÚÙÛÜ]/u'     =>   'U',
                '/ç/'           =>   'c',
                '/Ç/'           =>   'C',
                '/ñ/'           =>   'n',
                '/Ñ/'           =>   'N',
                '/–/'           =>   '-', // UTF-8 hyphen to "normal" hyphen
                '/[’‘‹›‚]/u'    =>   ' ', // Literally a single quote
                '/[“”«»„]/u'    =>   ' ', // Double quote
                '/ /'           =>   ' ', // nonbreaking space (equiv. to 0x160)
            );
            return preg_replace(array_keys($utf8), array_values($utf8), $text);
        }

        /* ============================== BEGIN PROSES ================================ */
        $sql_REG1 = "SELECT
            CONVERT(varchar,[Tgl_Tarik],20) [Tgl_Tarik],
            --[WilayahID],
            [WilayahName],
            [RegionID],
            --[OurBranchIDPNM],
            [Initial],
            --[BranchName],
            [RegionName],
            --[AreaID],
            [AreaName],
            [UnitID],
            [UnitName],
            [UnitInitial],
            --[Kota],
            [GroupID],
            [GroupName],
            [ClientID],
            [Name],
            [AccountID],
            [LoanSeries],
            [ProductID],
            [InstallmentAmount],
            FORMAT([FirstDisbursementDate], 'yyyy-MM-dd') [FirstDisbursementDate],
            [OutStandingPrincipal],
            [OutStandingInterest],
            [ODPrincipal],
            [ODInterest],
            [NoOfArrearDays],
            [DisbursedAmount],
            [Term],
            [MeetingDayID],
            [LoanTypeID],
            [RepaymentFrequencyID],
            [RepaymentTerm],
            FORMAT([LastPaidDate], 'yyyy-MM-dd') [LastPaidDate],
            [LastInstallmentNo],
            [PrincipalPaid],
            [InterestPaid],
            [UnearnPrincipal],
            [UnearnInterest],
            [PrincipalDue],
            [InterestDue],
            [InstallmentNo],
            FORMAT([InstallmentDueDate], 'yyyy-MM-dd') [InstallmentDueDate],
            [PaymentNo],
            [InterestAmount]
            FROM [<NAMA_DB>].[dbo].[<NAMA_TABLE>]    ";

        $stmt_REG1 = sqlsrv_query( $conn_SVR, $sql_REG1 );
        $count = 0;
        while( $row = sqlsrv_fetch_array( $stmt_REG1, SQLSRV_FETCH_ASSOC)) {        
            
            $string =  $row['Name'];
            $nama = cleanString($string);
            $Tgl_Tarik = $row['Tgl_Tarik'];
            $TGL =  date('Ymd', strtotime($Tgl_Tarik));
            $JAM = date('H', strtotime($Tgl_Tarik));
            $aa = iconv('UTF-8', 'ISO-8859-1//TRANSLIT//IGNORE', $string);

            $rows1[] = array(
                $row['Tgl_Tarik'],
                //$row['WilayahID'],
                $row['WilayahName'],
                $row['RegionID'],
                //$row['OurBranchIDPNM'],
                $row['Initial'],
                //$row['BranchName'],
                $row['RegionName'],
                //$row['AreaID'],
                $row['AreaName'],
                $row['UnitID'],
                $row['UnitName'],
                $row['UnitInitial'],
                //$row['Kota'],
                $row['GroupID'],
                $row['GroupName'],
                $row['ClientID'],
                $aa,
                $row['AccountID'],
                $row['LoanSeries'],
                $row['ProductID'],
                $row['InstallmentAmount'],
                $row['FirstDisbursementDate'],
                $row['OutStandingPrincipal'],
                $row['OutStandingInterest'],
                $row['ODPrincipal'],
                $row['ODInterest'],
                $row['NoOfArrearDays'],
                $row['DisbursedAmount'],
                $row['Term'],
                $row['MeetingDayID'],
                $row['LoanTypeID'],
                $row['RepaymentFrequencyID'],
                $row['RepaymentTerm'],
                $row['LastPaidDate'],
                $row['LastInstallmentNo'],
                $row['PrincipalPaid'],
                $row['InterestPaid'],
                $row['UnearnPrincipal'],
                $row['UnearnInterest'],
                $row['PrincipalDue'],
                $row['InterestDue'],
                $row['InstallmentNo'],
                $row['InstallmentDueDate'],
                $row['PaymentNo'],
                $row['InterestAmount'],
            
            );
            $count ++;
            echo "GET DATA ".$count." ROWS \n";
        }
    
        sqlsrv_free_stmt( $stmt_REG1);
        
        echo "TUNGGU BRO !! SEDANG DI CREATE FILE EXCEL NYA !! \n";
        echo "DATA LUMAYAN GEDE NIH JADI LAMA .. \n\n";

        $writer = new XLSXWriter();
        $writer->setAuthor('RKB'); 
        $writer->writeSheetHeader('NAMA_FILE_NYA', $header, $header_style);
        foreach($rows1 as $row)
        $writer->writeSheetRow('NAMA_FILE_NYA', $row);       
        $today = date('Ymd');
        //$jam = date("h");
        
        $filename = 'NAMA_FILE_NYA'.$TGL.'_Jam_'.$JAM.'.xlsx';
        $writer->writeToFile('FILE_EXCEL/'.$filename);

    
        echo "FILE EXCEL BERHASIL DI GENERATE \n";
        echo "SILAHKAN LIHAT FILE NYA DI FOLDER FILE_EXCEL \n\n";

        /* ============================== END PROSES  ================================ */


        /* ============================== PROSES INSERT LOG ================================ */
        echo "PROSES INSERT LOG\n\n";
       
        $TANGGAL_DB = date('Y-m-d');
        $TGL_TARIK = $TGL_TARIK;
        $REGION = 'ALL';
        $sql_insert = "INSERT INTO [LOG_JOB]
                    (
                        [TANGGAL_DB], [TGL_TARIK], [REGION], NAMA_FILE
                    )
                    VALUES
                    (
                        '$TANGGAL_DB', '$TGL_TARIK', '$REGION', '$filename' 
                    )";

        $stmt_insert = sqlsrv_query( $conn, $sql_insert );
        if( $stmt_insert === false) {
            die( print_r( sqlsrv_errors(), true) );
        }
        else{
            echo "PROSES INSERT LOG BERHASIL\n";
        }

        sqlsrv_free_stmt( $stmt_insert);
    }
   

    shell_exec('start cmd.exe @cmd /k "D:\CONSOLE_DATA\DISCONNECT_ANYCONNECT.bat" exit');

?>  