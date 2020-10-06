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

 async AddFile(){
  var reader = new window.FileReader();
  reader.onload = function (event) {
    // strip off the metadata before the base64-encoded string
    var startIndex = (event.target.result as string).indexOf("base64,");
    var copyBase64 = (event.target.result as string).substr(startIndex + 7);
    Office.context.document.setSelectedDataAsync(copyBase64);    
  };
 }

}
