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

  async goToFirstSlide() {
    var goToFirst = Office.Index.First;

    Office.context.document.goToByIdAsync(goToFirst, Office.GoToType.Index);
  }

  async goToLastSlide() {
    var goToLast = Office.Index.Last;

    Office.context.document.goToByIdAsync(goToLast, Office.GoToType.Index);
  }

  async goToNextSlide() {
    var goToNext = Office.Index.Next;

    Office.context.document.goToByIdAsync(goToNext, Office.GoToType.Index);
    this.countNext++;
  }

  async goToPrevSlide() {
    var goToPrev = Office.Index.Previous;

    Office.context.document.goToByIdAsync(goToPrev, Office.GoToType.Index);
    this.countPrev++;
  }

  async changeColor(){
   var random = Math.floor(Math.random() * 2);
   return random == 0 ? this.color = "green" : this.color = "red";
  }
}
