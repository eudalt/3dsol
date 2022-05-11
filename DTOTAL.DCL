DATOS :dialog {
  label = "EXPERTLINE ";
   : row  {
   : button {
        key = "FICHERO";
        label = "ARXIU DADES...";
   }

   : button {
        key = "BUSCA";
        label = "NOMES TROBATS";
   }
   }
 : row  {	    : edit_box {
	        label = "Altura";
	        key = "ALT";
	        width = 10;
	        alignment = right;
	      }
   : button {
        key = "BORRAL";
        label = "BORRA LINIA";
   }
   }
    : text {
      key = "TEXT1";
      }   
   :list_box {
       key = "VLISTA";
       label = " QUANTITAT X ALTURA  RESULTAT  DESCRIPCIó ";
       width = 60;
       height = 25;
      }
     

	      : row  {

     : edit_box {
        label = "NOM ARXIU A GUARDAR: ";
        key = "NOMDAT";
        width = 10;
        alignment = right;
      }: button {
        key = "SALVA4";
        label = "SALVA EXCEL 'XLS'";
   }
   }
      : row  {
       : button {
        key = "BLOCS";
        label = "BUSCA BLOCS";
   }
   ok_only;
   }
}  

POLPAR :dialog {
 label = "PARAMETROS";
      : edit_box {
        label = "NOMBRE VARIABLE";
        key = "VAR";
        width = 10;
        alignment = right;
      }
      : edit_box {
        label = "VALOR REAL DEL BLOQUE";
        key = "VAL";
        width = 10;
        alignment = right;
      }
      : edit_box {
        label = "DIVISOR DEL TOTAL";
        key = "DIV";
        width = 10;
        alignment = right;
      }
  
}
