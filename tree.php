<?php
$filename = "page1.json";

         $file = fopen( $filename, "r" );

         if( $file == false ) {
            echo ( "Error in opening file" );
            exit();
         }

         $filesize = filesize( $filename );

         $filetext = fread( $file, $filesize );
         fclose( $file );

         echo ($filetext);

?>
