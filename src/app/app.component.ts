import { Component } from '@angular/core';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent {
  async AddFile(input: HTMLInputElement) {
    var file: File = input.files[0];
    var reader = new window.FileReader();
    reader.readAsDataURL(file);
    reader.onload = () => {
      var startIndex = (reader.result as string).indexOf('base64,');
      var copyBase64 = (reader.result as string).substr(startIndex + 7);
      Office.context.document.setSelectedDataAsync(
        copyBase64,
        {
          coercionType: Office.CoercionType.Image
        },
        function(asyncResult) {
          if (String(asyncResult.status) == 'failed') {
            console.log('Error: ' + asyncResult.error.message);
          }
        }
      );
    };
  }
}
