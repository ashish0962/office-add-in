# office-add-in

### Installation

_Below is an example of how you can install and setup your app._
1. Clone the repo
   ```sh
   git clone https://github.com/ashish0962/office-add-in.git
   ```
2. Install NPM packages
   ```sh
   npm install
   ```

### Test the add-in

* To test your add-in in Excel, run the following command in the root directory of your project. This starts the local web server (if it's not already running) and opens Excel with your add-in loaded.
```sh
npm start
```
* If you're testing your add-in on Mac, run the following command in the root directory of your project before proceeding. When you run this command, the local web server starts.
```sh
npm run dev-server
```
* To test your add-in in Excel on the web, run the following command in the root directory of your project. When you run this command, the local web server starts. Replace "{url}" with the URL of an Excel document on your OneDrive or a SharePoint library to which you have permissions.
```sh
npm run start:web -- --document {url}
```
  If your add-in doesn't sideload in the document, manually sideload it by following the instructions in [Sideload Office Add-ins in Office on the web manually](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing#sideload-an-office-add-in-in-office-on-the-web-manually).
