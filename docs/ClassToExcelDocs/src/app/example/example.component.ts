import { Component, OnInit, Input } from '@angular/core';
import { ExampleService } from './example.service';

@Component({
  selector: 'cte-example',
  templateUrl: './example.component.html',
  styleUrls: ['./example.component.css']
})
export class ExampleComponent implements OnInit {

  @Input() fileName: string = "hello";
  exampleCode: string = "nada here yet!";

  constructor(private _exampleService: ExampleService) { }

  ngOnInit() {
    this._exampleService.getExample(this.fileName)
      .subscribe(data => this.exampleCode = this.simpleFormat(data));
  }

  simpleFormat(data: string): string {
    let result: string = "";
    for (let i = 0; i < data.length; i++) {
      switch (data[i]) {
        case "<":
          result += "&lt;"
          break;
          case ">":
          result += "&gt;"
          break;
        default:
          result += data[i];
          break;
      }
    }

    return result;
  }
}
