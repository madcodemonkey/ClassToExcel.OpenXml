import { Component, OnInit } from '@angular/core';
import { ExampleService } from '../../example.service';

@Component({
  templateUrl: './class-to-excel-raw-reader.component.html'
})
export class ClassToExcelRawReaderComponent implements OnInit {

  constructor(private _exampleService: ExampleService) { }

  example1CodePart1: string;
  example1CodePart2: string;
  example2Code: string;

  ngOnInit() {
    this._exampleService.getExample("RawReaderExample1Part1.txt")
      .subscribe(data => this.example1CodePart1 = data);
    this._exampleService.getExample("RawReaderExample1Part2.txt")
      .subscribe((data) => this.example1CodePart2 = data);
    this._exampleService.getExample("RawReaderExample2.txt")
      .subscribe((data) => this.example2Code = data);
  }

}
