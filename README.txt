Hello, I made these programs to help speed up bulk returns of any warehouse customer, but these programs are already pre-made and optimized for Cabana UPC to SKU conversion. 

Follow these steps to run them and to help speed up the return process of all the loose items.

1. Run SKU_Match.py
    A. Insert the name of the stock status export that you are going to use.  
        EX: For Cabana you would type the name of the excel file in this folder: Cabana_INV.xlsx
    B. Next you will insert the name of the excel file that you will save the compared sku's to.
        EX: I liked to name the file: Found.xlsx
    C. Start scanning your items from the given UPC or barcode.
    D. If you finish or want to take a break just type "exit".
    E. After exiting the program will make an excel file called: SKUS.xlsx that
        contains all the skus found and has them all added up. I like to check if any are missing.
    F. Check the file you made in part (B) and see if the sku column has any empty 
        spaces. If so that means the sku was not found and you can try looking it
        up on Extensiv. If it does not appear just copy what you scanned into that
        empty space and let Cabana deal with it later.



    
2. Most of the time Cabana will email you a list o QVC items being sent back or they will have a RMA already made. 
A big portion of the time these RMA's do not match what we have received in our list.
