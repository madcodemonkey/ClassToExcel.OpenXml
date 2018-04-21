import { Component, OnInit } from '@angular/core';
import { ExampleService } from '../../example.service';

@Component({
  templateUrl: './class-to-excel-reader.component.html'
})
export class ClassToExcelReaderComponent implements OnInit {

  constructor(private _exampleService: ExampleService) { }

  step1Code: string;
  step2Code: string;
  step3Code: string;

  ngOnInit() {
    this._exampleService.getExample("ReaderStep1.txt")
      .subscribe(data => this.step1Code = data);
    this._exampleService.getExample("ReaderStep2.txt")
      .subscribe(data => this.step2Code = data);
    this._exampleService.getExample("ReaderStep3.txt")
      .subscribe(data => this.step3Code = data);
  }

}
