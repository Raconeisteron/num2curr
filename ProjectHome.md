num2curr Excel-Addin v1.1

num2curr addin has 2 excel functions

1) num2curr() -> converts numbers to currency(Rupees) rounding upto 2 decimal places.

2) num2word() -> converts numbers to words

Features:

-> Supports both indian and international system.

-> Auto detects the number system fom the comma pattern of the number.

-> Easy to install and unistall

To install the Excel-function:

1) Download the file from http://num2curr.googlecode.com/files/num2curr%20v1.1.zip

2) extract it

3) Close Excel if already running

4) Then double-click num2curr.exe and follow the instructions.

To uninstall the Excel-function:

1) Double click num2curr.exe then click yes when prompted for uninstallation


Usage:

After installing the program open a  Excel File

Now you can use the num2curr or num2word function by giving arguments (number,number\_system) comma pattern is ignored if number\_system argument is given.

This can be used as any other normal function in excel

Examples:

1) To convert 123456.123 to currency in indian system(lakhs, crores)

> Input: =num2curr(123456.123) (or) =num2curr("1,23,456.123") (or) =num2curr(123456.123,"indian") (or) =num2curr(123456.123,"ind")

> Result: One Lakh Twenty Three Thousand Four Hundred And Fifty Six Rupees Twelve Paise Only


2) To convert 123456.123 into currency in international system(millions,billions,trillions)

> Input: =num2curr("123,456.123") (or) =num2curr(123456.123,"international") (or) =num2curr(123456.123,"int")

> Result: One Hundred And Twenty Three Thousand Four Hundred And Fifty Six Rupees Twelve Paise Only


3) To convert numbers to words(doesn't append Ruppes and doesn't round off fractional part) use num2word instead.

> Input: =num2word(123456.123) (or) =num2word("1,23,456.123") (or) =num2word(123456.123,"indian") (or) =num2word(123456.123,"ind")

> Result: One Lakh Twenty Three Thousand Four Hundred And Fifty Six Point One Two Three


For Queries and reporting any bugs:
> Contact M.SainathGupta (m.sainathgupta@gmail.com)