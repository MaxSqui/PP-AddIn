import { NgModule } from "@angular/core";
import { BrowserModule } from "@angular/platform-browser";
import AppComponent from "./app.component";
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';

const timer = {};

function startInterval(timer) {
  timer.id = setInterval(() => {
    if (Office !== undefined) {
      Office.initialize = (reason) => {
        platformBrowserDynamic().bootstrapModule(AppModule).catch(error => console.error(error));
      };
      stopInterval(timer.id);
    }
  }, 1000);
}

function stopInterval(id) {
  clearInterval(id);
}

startInterval(timer);

@NgModule({
  declarations: [AppComponent],
  imports: [BrowserModule],
  bootstrap: [AppComponent]
})
export default class AppModule {}
