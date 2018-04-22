import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { Routes, RouterModule } from '@angular/router';

import { ReadingComponent } from './reading.component';
import { ClassToExcelReaderComponent } from './readers/class-to-excel-reader.component';
import { ClassToExcelRawReaderComponent } from './readers/class-to-excel-raw-reader.component';
import { ExampleModule } from '../example/example.module';
import { ClassToExcelRowReaderComponent } from './readers/class-to-excel-row-reader.component';

const routes: Routes = [{
    path: "reading",
    children: [
      { path: "", component: ReadingComponent},
      { path: "ClassToExcelReader", component: ClassToExcelReaderComponent },
      { path: "ClassToExcelRawReader", component: ClassToExcelRawReaderComponent },
      { path: "ClassToExcelRowReader", component: ClassToExcelRowReaderComponent}
    ]
}];

@NgModule({
  imports: [
    RouterModule.forChild(routes),
    CommonModule,
    ExampleModule
  ],
  declarations: [
    ReadingComponent,
    ClassToExcelReaderComponent,
    ClassToExcelRawReaderComponent,
    ClassToExcelRowReaderComponent
  ]
})
export class ReadingModule { }
