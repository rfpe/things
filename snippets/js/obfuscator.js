/*
This small algorithm will help you to hide your e-mail address from e-mail crawler bots.
There's an example .xlsx file on this directory (example.xlsx)

  How to use
 
 1) Create your string. I.E. <a href='mailto:some@email.com'>some@email.com</a>
 2) Create a new Excel spreadsheet
 3) The columns will be:
	3.1) A= Index (from 0 to the lenght of your string)
	3.2) B= Characters of your string (place on character per row)
	3.2) C= Use the function Random() for each line
 4) Add Autofilter 
 5) Sort ascending by Column C ("Random")
 6) Copy the sorted content of column A, make it an array (first_indexes)
 7) Copy the sorted content of column B, make it an array (characters_array) 

 Deploy
 
 1) Change the variables' names to gibberish
 2) Place this script between <script type="text/javascript"> and </script> tags.
 3) There you go! 
 
 This algorithm was firstly developed by Paulo Higa (http://higa.me)
 

*/

var first_indexes = [
47,3,32,33,34,43,6,21,2,31,
35,40,25,9,48,10,15,18,13,38,
12,20,29,28,42,5,27,46,26,44,
16,45,23,7,4,1,14,36,49,41,
11,24,19,30,22,0,17,8,37,39];

var characters_array = [
'/','h','s','o','m','c','f','e',' ','>',
'e','i','l','m','a','a',':','m','t','m',
'l','@','m','o','.','e','c','<','.','o',
's','m','a','=','r','a','o','@','>','l',
'i','i','e',' ','m','<','o',' ','e','a'];

var output_array = Array();

/*
First iteraction:
	output_array[first_indexes[0]] = characters_array[0];
	output_array[47] = '/'

Second iteraction:
	output_array[first_indexes[1]] = characters_array[1];
	output_array[3] = 'h'

	and so on...
*/
for(var i=0;i<first_indexes.length;i++){
	output_array[first_indexes[i]] = characters_array[i];
}

//Writes each characters on output_array to screen
for(var j=0;j<output_array.length;j++){
	document.write(output_array[j]);
}