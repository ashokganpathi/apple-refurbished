# apple-refurbished
Pull information from Apple's refurbished products on sale and create a per day comparison price sheet.


The code is an excel macro that fetches the current day apple refurbished product prices from a set of URLs in the configuration file.

Then it collects this information in an output sheet with the list of products and their current day's price.
The prices are color coded as follows - 

For a given product, 
 if the price from today is the same as the previous run, the font color is black.
 if the price from today is less than the previous price, the font color is green.
 if the price from today is greater than the previous price, the font color is red.

If a given product is not available on sale on a given day, the cell is blank and no price is recorded.
