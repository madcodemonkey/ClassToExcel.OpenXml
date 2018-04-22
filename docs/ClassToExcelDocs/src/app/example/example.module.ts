import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { HttpClientModule } from '@angular/common/http';

import { ExampleComponent } from './example.component';
import { ExampleService } from './example.service';

@NgModule({
  imports: [
    CommonModule,
    HttpClientModule
  ],
  exports:[
    ExampleComponent
  ],
  declarations: [ExampleComponent],
  providers: [ExampleService]
})
export class ExampleModule { }
