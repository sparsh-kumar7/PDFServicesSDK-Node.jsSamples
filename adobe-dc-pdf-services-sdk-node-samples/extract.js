const PDFServicesSdk = require('@adobe/pdfservices-node-sdk');
const fs = require('fs');
const AdmZip = require('adm-zip');


try {
    // Initial setup, create credentials instance.
    const credentials =  PDFServicesSdk.Credentials
        .servicePrincipalCredentialsBuilder()
        .withClientId("323ea124415c4c1699cc8032abaf2dbc")
        .withClientSecret("p8e-8FQwInVnMYUwSVw69fjmuaamLGvqEQ_a")
        .build();

    // Create an ExecutionContext using credentials
    const executionContext = PDFServicesSdk.ExecutionContext.create(credentials);

    // Build extractPDF options
    const options = new PDFServicesSdk.ExtractPDF.options.ExtractPdfOptions.Builder()
        .addElementsToExtract(PDFServicesSdk.ExtractPDF.options.ExtractElementType.TEXT, PDFServicesSdk.ExtractPDF.options.ExtractElementType.TABLES)
        .addElementsToExtractRenditions(PDFServicesSdk.ExtractPDF.options.ExtractRenditionsElementType.FIGURES, PDFServicesSdk.ExtractPDF.options.ExtractRenditionsElementType.TABLES)
        .build();

    // Create a new operation instance.
    const extractPDFOperation = PDFServicesSdk.ExtractPDF.Operation.createNew(),
        input = PDFServicesSdk.FileRef.createFromLocalFile(
            'resources/extractPDFInput.pdf',
            PDFServicesSdk.ExtractPDF.SupportedSourceFormat.pdf
        );

    // Set operation input from a source file
    extractPDFOperation.setInput(input);

    // Set options
    extractPDFOperation.setOptions(options);

    extractPDFOperation.execute(executionContext)
        .then(result => result.saveAsFile('output/ExtractTextTableWithFigureTableRendition.zip'))
        .catch(err => {
            if(err instanceof PDFServicesSdk.Error.ServiceApiError
                || err instanceof PDFServicesSdk.Error.ServiceUsageError) {
                console.log('Exception encountered while executing operation', err);
            } else {
                console.log('Exception encountered while executing operation', err);
            }
        });
} catch (err) {
    console.log('Exception encountered while executing operation', err);
}   


const unzipper = require('unzipper');
const createCsvWriter = require('csv-writer').createObjectCsvWriter;

try {
  // Initial setup, create credentials instance.
  const credentials = PDFServicesSdk.Credentials
    .servicePrincipalCredentialsBuilder()
    .withClientId("PDF_SERVICES_CLIENT_ID")
    .withClientSecret("PDF_SERVICES_CLIENT_SECRET")
    .build();

  // Create an ExecutionContext using credentials
  const executionContext = PDFServicesSdk.ExecutionContext.create(credentials);

  // Build extractPDF options
  const options = new PDFServicesSdk.ExtractPDF.options.ExtractPdfOptions.Builder()
    .addElementsToExtract(PDFServicesSdk.ExtractPDF.options.ExtractElementType.TEXT, PDFServicesSdk.ExtractPDF.options.ExtractElementType.TABLES)
    .addElementsToExtractRenditions(PDFServicesSdk.ExtractPDF.options.ExtractRenditionsElementType.TABLES)
    .addTableStructureFormat(PDFServicesSdk.ExtractPDF.options.TableStructureType.CSV)
    .build();

  // Create a new operation instance.
  const extractPDFOperation = PDFServicesSdk.ExtractPDF.Operation.createNew(),
    input = PDFServicesSdk.FileRef.createFromLocalFile(
      'resources/extractPDFInput.pdf',
      PDFServicesSdk.ExtractPDF.SupportedSourceFormat.pdf
    );

  // Set operation input from a source file.
  extractPDFOperation.setInput(input);

  // Set options
  extractPDFOperation.setOptions(options);

  extractPDFOperation.execute(executionContext)
    .then(result => result.saveAsFile('output/ExtractTextTableWithTableStructure.zip'))
    .then(() => {
      // Extract the JSON file from the ZIP archive
      return unzipper.Open.file('output/ExtractTextTableWithTableStructure.zip');
    })
    .then((archive) => {
      const jsonEntry = archive.files.find(file => file.path.endsWith('.json'));
      if (jsonEntry) {
        return jsonEntry.buffer();
      } else {
        throw new Error('JSON file not found in the archive.');
      }
    })
    .then((buffer) => {
      // Parse the JSON data
      const jsonData = JSON.parse(buffer.toString());

      // Define the column names as keys in an object
      const columnNames = {
        'Bussiness__City': 'Bussiness::City',
        'Bussiness__Country': 'Bussiness::Country',
        'Bussiness__Description': 'Bussiness::Description',
        'Bussiness__Name': 'Bussiness::Name',
        'Bussiness__StreetAddress': 'Bussiness::StreetAddress',
        'Bussiness__Zipcode': 'Bussiness::Zipcode',
        'Customer__Address__line1': 'Customer::Address::line1',
        'Customer__Address__line2': 'Customer::Address::line2',
        'Customer__Email': 'Customer::Email',
        'Customer__Name': 'Customer::Name',
        'Customer__PhoneNumber': 'Customer::PhoneNumber',
        'Invoice__BillDetails__Name': 'Invoice::BillDetails::Name',
        'Invoice__BillDetails__Quantity': 'Invoice::BillDetails::Quantity',
        'Invoice__BillDetails__Rate': 'Invoice::BillDetails::Rate',
        'Invoice__Description': 'Invoice::Description',
        'Invoice__DueDate': 'Invoice::DueDate',
        'Invoice__IssueDate': 'Invoice::IssueDate',
        'Invoice__Number': 'Invoice::Number',
        'Invoice__Tax': 'Invoice::Tax'
      };

      // Extract the data points against the defined column headings
      const records = [columnNames]; // Add the column names as the first record
      const dataRecord = {};

      // Iterate through the column names and find their corresponding values in the JSON data
      for (const columnName in columnNames) {
        const [section, ...keys] = columnNames[columnName].split('__');
        let value = jsonData;

        // Traverse the keys to get the corresponding value
        for (const key of keys) {
          if (value[section]) {
            value = value[section][key];
          } else {
            value = null;
            break;
          }
        }

        dataRecord[columnName] = value;
      }

      // Add the extracted data record to the records array
      records.push(dataRecord);

      // Create CSV writer and write the records to the CSV file
      const csvWriter = createCsvWriter({
        path: 'output/ExtractedData.csv',
        header: Object.keys(columnNames).map(column => ({ id: column, title: column }))
      });

      return csvWriter.writeRecords(records);
    })
    .then(() => {
      console.log('CSV file generated successfully.');
    })
    .catch(err => {
      console.error('Exception encountered while executing operation', err);
    });
} catch (err) {
  console.error('Exception encountered while executing operation', err);
}
