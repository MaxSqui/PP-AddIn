import { Component } from "@angular/core";

// images references in the manifest
import "../../../assets/icon-16.png";
import "../../../assets/icon-32.png";
import "../../../assets/icon-80.png";
const template = require("./app.component.html");
/* global console, Office, require */
/* global PowerPoint*/

@Component({
  selector: "app-home",
  template
})
export default class AppComponent {

 async AddFile(input: HTMLInputElement) {
  var file: File = input.files[0];
  var reader = new window.FileReader();
  reader.readAsDataURL(file);
  reader.onload = () => {
    var startIndex = (reader.result as string).indexOf("base64,");
     var copyBase64 = (reader.result as string).substr(startIndex + 7);
     Office.context.document.setSelectedDataAsync(copyBase64,
      {
       coercionType: Office.CoercionType.Text
      },
      function (asyncResult) {
        if (String(asyncResult.status) == "failed") {
            console.log('Error: ' + asyncResult.error.message);
        }
      });
    };
  } 
}
