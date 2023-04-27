# Templator
Templator is a simple console application that reads a Word document and replaces placeholders with values from an Excel spreadsheet. It also has an optional ability to generate PDF files by using libreoffice.

## Usage
Add some placeholders to a word document using hashtags, eg.: Hello, **#name#**!.
Then, in the spreadsheet, add a column with the same name as the placeholder, in this example it should be **name**.  
When running Templator, you'll need to provide:

    Template path
    Spreadsheet path
    Destination folder path

Templator will generate a word document for each line in the spreadsheet, replacing the placeholders with the corresponding data from the spreadsheet.

## Acknowledgments
This project uses the following open-source libraries:

- [DocumentFormat.OpenXml](https://github.com/OfficeDev/Open-XML-SDK)
- [Spectre.Console](https://github.com/spectresystems/spectre.console)

## License
This project is licensed under the MIT License - see the LICENSE.md file for details.