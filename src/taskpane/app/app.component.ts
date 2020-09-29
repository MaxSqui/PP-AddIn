import { Component } from "@angular/core";
import { P } from "@angular/core/src/render3";
// images references in the manifest
import "../../../assets/icon-16.png";
import "../../../assets/icon-32.png";
import "../../../assets/icon-80.png";
import pptxgen from "pptxgenjs";
const template = require("./app.component.html");
/* global console, Office, require */

@Component({
  selector: "app-home",
  template
})
export default class AppComponent {
  welcomeMessage = "Hi!";
  color : string;

  countNext: number = 0;
  countPrev: number = 0;

  goToNextSlide() {
    var goToNext = Office.Index.Next;

    Office.context.document.goToByIdAsync(goToNext, Office.GoToType.Index);
    this.countNext++;
}

goToPrevSlide() {
  var goToPrev = Office.Index.Previous;

  Office.context.document.goToByIdAsync(goToPrev, Office.GoToType.Index);
  this.countPrev++;
}

  async run() {
    /**
     * Insert your PowerPoint code here
     */
    Office.context.document.setSelectedDataAsync(
      "Hello World!",
      {
        coercionType: Office.CoercionType.Text
      },
      result => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error(result.error.message);
        }
      }
    );
  }

  async changeColor(){
   var random = Math.floor(Math.random() * 2);
   return random == 0 ? this.color = "green" : this.color = "red";
  }
}
