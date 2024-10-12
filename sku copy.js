const xlsx = require('xlsx');
const fs = require('fs');
const readline = require('readline');

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});



// get the args
const args = process.argv.slice(2);
const file = 1

// Load the XLSX file
const workbook = xlsx.readFile('STOCK FINAL.xlsx');
const sheetName = workbook.SheetNames[Number(file)];
const worksheet = workbook.Sheets[sheetName];

// Convert the sheet to JSON for easier manipulation
let productsFromExcel = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

const askQuestion = (query) => {
  return new Promise((resolve) => rl.question(query, resolve));
};

const products = [];
let currentHeader = null;
let currentProduct = null;
let variants = [];

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
let productName = "";
const productToSku = []

// Process products and format them into two columns: SKU and variant values
processProducts(productsFromExcel).then(() => {
  


  const names = Object.keys(products[0].values).filter(k => k.includes("REF")).map(k => k.split("REF ")[1]);
  const obj = {}
  names.forEach(name => {
    products.forEach(product => {
        obj[product.values[name]] = []
    })
  });
  names.forEach(name => {
    products.forEach(product => {
      if(!obj[product.values[name]].includes(product.values["REF " + name])){
        obj[product.values[name]].push(product.values["REF " + name])
      }
    })
  });

  // console.log(obj);
  const newObj = {}
  Object.keys(obj).forEach(key => {
    if(obj[key].length > 1) {
       newObj[key] = obj[key]
    }
  })
  console.log(Object.keys(newObj));
  // fs.writeFileSync("refs_filtered.json", JSON.stringify(newObj, null, 2));
  // fs.writeFileSync("refs_names.json", JSON.stringify(
  //   Object.keys(newObj)
  // , null, 2));

// (
//   products.find(p=>p.values["FABRICANT"]) ?
//   products.filter(p=>["GENERIQUE","AKSES"].includes(p.values["FABRICANT"])):products
// )
//   .forEach((product) => {
//     if(productName !== product.name) {
//       productName = product.name;
//       productToSku.push(
//         {
//           "Product Name":productName,
//           "SKU Forma": Object.keys(
//             product.values
//           ).filter(key => key.startsWith("REF ")).map(key => key.replace("REF ","")).join(" - "),
//           "Example SKU": Object.keys(
//             product.values
//           ).filter(key => key.startsWith("REF ")).map(key =>{
//             const variant = product.values[key]
//             if(key=="GAMME" && !variant ) {
//               return "PREMIUM"
//             } else {
//               return variant ?? "_"
//             }
//           } ).join("-"),
//         }
//       );
//     }
// // 


//     // Get the SKU (either "REF FABRICANT" or any other "REF " key)
//     // const sku = product.values["REF FABRICANT"] ? 
//     //   "---":
//     //   Object.keys(product.values).filter(key => key.startsWith("REF ")).map(key => product.values[key]).join("-");

//     // // Get the variant values as a string
//     // const variantValues = product.variants.map(v => {
//     //   if (v === "GAMME") {
//     //     return product.values[v] ?? "PREMIUM";
//     //   } else {
//     //     return product.values[v] ?? "_";
//     //   }
//     // }).join("#");

//     // Add the SKU and variant values to the reformatted products array
//     // if(sku !== "---") {
//     //   reformattedProducts.push({
//     //     "Variant Values": variantValues,
//     //     SKU: sku,
//     //   });
//     // }
  // });

  // console.log(productToSku);
  // // Create a new XLSX file with the reformatted products
  // const newWorksheet = xlsx.utils.json_to_sheet(reformattedProducts);
  // const newWorkbook = xlsx.utils.book_new();
  // xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, 'SKU and Variants');

  // // Write the file as 'sku_variants.xlsx'
  // xlsx.writeFile(newWorkbook, 'sku_variants.xlsx');

  // console.log('File generated successfully as sku_variants.xlsx');
  // console.log(productToSku);
  // const worksheet = xlsx.utils.json_to_sheet(productToSku);
  // const workbook = xlsx.utils.book_new();
  // xlsx.utils.book_append_sheet(workbook, worksheet, 'SKU and Variants');
  // xlsx.writeFile(workbook, './universes/'+sheetName+'_sku_variants_.xlsx');

  console.log('File generated successfully as sku_variants.xlsx');
});
