# apple-refurbished
Pull information from Apple's refurbished products on sale and create a per day comparison price sheet.
This is a personal utility that I have used to review daily price changes on Apple products in the refurbished section of the site.


The code is an excel macro that fetches the current day apple refurbished product prices from a set of URLs in the configuration file.

Then it collects this information in an output sheet with the list of products and their current day's price.
The prices are color coded as follows - 

For a given product, 
 if the price from today is the same as the previous run, the font color is black.
 if the price from today is less than the previous price, the font color is green.
 if the price from today is greater than the previous price, the font color is red.

If a given product is not available on sale on a given day, the cell is blank and no price is recorded.


**How to use the tool?**

1. Download the get-apple-refurbished-info.xlsm, apple-refurbished-info.xlsx [optional], apple-deals-urls.txt onto a folder in your PC.

2. Create a logs folder to ensure that logs are written in the folder without any errors.

3. Open the .txt file and add/remove any URLs for products that you're interested in.

![alt text](https://github.com/ashokganpathi/apple-refurbished/blob/master/img/01-apple-deals-urls.png)

4. Open the .xlsm file and click on "Get Deals" button. Once it's completed, it would give a pop-up indicating that the macro has completed running.

![alt text](https://github.com/ashokganpathi/apple-refurbished/blob/master/img/02-run-macro.png)

5. Close the pop-up and the .xlsm file and open the apple-refurbished-info.xlsx file. If you did not download the existing file, it would have created a new file with today's list of products. If you had downloaded the file, it would have appended a new column with today's date and put in the product prices along w/ color coding based on comparison with the last run.

![alt text](https://github.com/ashokganpathi/apple-refurbished/blob/master/img/03-output-sheet.png)

Hope you have fun tweaking or extending the utility!
