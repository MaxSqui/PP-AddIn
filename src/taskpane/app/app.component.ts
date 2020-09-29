import { Component } from "@angular/core";
// images references in the manifest
import "../../../assets/icon-16.png";
import "../../../assets/icon-32.png";
import "../../../assets/icon-80.png";
const template = require("./app.component.html");
/* global console, Office, require */

@Component({
  selector: "app-home",
  template
})
export default class AppComponent {
  welcomeMessage = "Hi!";
  color : string;

  countNext: number ;
  countPrev: number ;

  currSlide:number;


  async addEventHandlerToBinding() {
    Office.context.document.addHandlerAsync(
        Office.EventType.ActiveViewChanged, this.onBindingSlideChanged);
  }

async onBindingSlideChanged() {
        var lastSlide;
        if(this.currSlide!=Office.Index.Next-1){
          lastSlide = this.currSlide;
          this.currSlide = Office.Index.Next-1;
        }
        if(this.currSlide<lastSlide){
          this.countPrev++;
        }
        else {
          this.countNext++;
        }
        lastSlide=this.currSlide;
}


  async run() {
    Office.EventType.ActiveViewChanged
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
