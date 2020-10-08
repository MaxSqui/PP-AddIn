import { enableProdMode } from '@angular/core';
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';

import { AppModule } from './app/app.module';
import { environment } from './environments/environment';

if (environment.production) {
  enableProdMode();
}

Office.initialize = () => {
  platformBrowserDynamic()
    .bootstrapModule(AppModule)
    .catch(err => displayError(err));
};
// Office.onReady(() => {
//   try {
//     throw Error('heelo');
//   } catch (error) {
//     displayError(error);
//   }
// });
function displayError(err: Error) {
  document.getElementById('error').innerText =
    err.name + ' ' + err.message + ' ' + err.name + ' ' + err.stack;
}
