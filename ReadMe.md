# Amazon Product Scraper

## Description
This is a simple .NET console application for scraping Amazon product information. It uses the HtmlAgilityPack and EPPlus packages to extract product title, price, and image URL from Amazon product pages. The application generates an Excel file containing the scraped product information.

## Features
- Scrapes product title, price, and image URL from Amazon product pages.
- Generates an Excel file with the scraped product information.

## Installation
To run the application, make sure you have .NET SDK installed on your machine.

```bash
# Clone the repository


# Navigate to the project directory
cd 

# Restore the NuGet packages
dotnet restore

# Build the project
dotnet build

# Run the application
dotnet run
```


## Usage
- When prompted, enter the desired product information including Product Name, Product Number, and Category.
- The application will scrape the Amazon product page for the specified product and generate an Excel file with the scraped product information.

## Dependencies
- HtmlAgilityPack (Version 1.11.59)
- EPPlus (Version 7.0.9)

## Note
- This application is intended for personal use and educational purposes only. Use responsibly and respect the terms of service of the websites you scrape.

## License
This project is licensed under the MIT License - see the LICENSE file for details.