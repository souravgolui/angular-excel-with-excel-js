import { Component, VERSION } from '@angular/core';
import { ExcelService } from './excel.service';

@Component({
  selector: 'my-app',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  name = 'Angular ' + VERSION.major;

  constructor(private excelService: ExcelService) {}

  downloadExcel() {
    this.excelService.generateExcel(this.name);
  }
}
