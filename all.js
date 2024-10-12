const xlsx = require('xlsx');
const fs = require('fs');
const readline = require('readline');

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

// Load the XLSX file
const workbook = xlsx.readFile('STOCK FINAL.xlsx');

// Loop through the sheets 1 to 8
const sheets = workbook.SheetNames.slice(1, 10);  // Get the first 8 sheets

const askQuestion = (query) => {
  return new Promise((resolve) => rl.question(query, resolve));
};

const products = [];
let currentHeader = null;
let currentProduct = null;
let variants = [];

// Function to split an array into chunks of specified size
const chunkArray = (array, chunkSize) => {
  const chunks = [];
  for (let i = 0; i < array.length; i += chunkSize) {
    chunks.push(array.slice(i, chunkSize + i));
  }
  return chunks;
};

// This needs to be an async function to use await
async function processProducts(productsFromExcel) {
  for (const line of productsFromExcel) {
    if (line[0] == "Image") {
      currentHeader = line;
      const getVariant = line.filter(l => l.includes("REF "));
      variants = getVariant.map(v => v.split("REF ")[1]);
      continue; // Move to the next line
    }

    if (line.length == 1) {
      console.log(line[0]);
      // This is a product name
      currentProduct = line[0];
      variants = []; // Reset variants for the new product
      continue;
    }

    if (line && line.length > 1) {
      // This is a product with headers, variants, etc.
      products.push({
        name: currentProduct,
        headers: currentHeader,
        values: Object.fromEntries(
          line
            .map((e, i) => {
              // Return an entry only if not empty
              if (e) return [currentHeader[i], e];
            })
            .filter(Boolean) // Filter out any undefined entries
        ),
        variants: variants
      });
    }
  }

  // Once done, close the readline interface
  rl.close();
}

let reformattedProducts = [];
let UniqueIdentification = 999;
let productName = "";

// Function to write each chunk of products to an XLSX file
function writeChunkToFile(chunk, index) {
  const newWorksheet = xlsx.utils.json_to_sheet(chunk);
  const newWorkbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, 'Reformatted Products');
  xlsx.writeFile(newWorkbook, `./odoo_products_import/part_${index + 1}.xlsx`);
}

// Function to process all sheets and combine them
async function processAllSheets() {
  for (const sheetName of sheets) {
    const worksheet = workbook.Sheets[sheetName];

    // Convert the sheet to JSON for easier manipulation
    let productsFromExcel = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
    
    await processProducts(productsFromExcel);
  }

  // Process and reformat the products
  products.forEach((product, productIndex) => {
    if (productName != product.name) {
      productName = product.name;
      UniqueIdentification = UniqueIdentification + 1;
    }

    reformattedProducts.push({
      "PRODUCT ID": product.values["PRODUCT ID"] ?? "no id",
      Name: product.name,
      "variant Attributes": product.variants.join(","),
      "Attribute Values": product.variants.map(v => {
        if(v === "GAMME") {
          return product.values[v] ?? "PREMIUM"; // highlight "Premium" in green
        } else {
          return product.values[v] ?? "_";
        }
      }).join("#"),
      "Internal Reference": 
        product.values["REF FABRICANT"] ? 
        product.values["REF FABRICANT"] :
        Object.keys(product.values).filter(key => key.startsWith("REF ")).map(key => product.values[key]).join("-"),
      "Category": product.values["CATEGORY"] ? product.values["CATEGORY"] : "All",
      "Image path/url": product.values.Image,
    });
  });

  // Split products into chunks (50,000 in each chunk)
  const productChunks = chunkArray(reformattedProducts, 500);

  // Write each chunk to a separate XLSX file
  productChunks.forEach((chunk, index) => {
    writeChunkToFile(chunk, index);
  });

  console.log('Files generated successfully.');
}

// Start the processing of all sheets
processAllSheets();
