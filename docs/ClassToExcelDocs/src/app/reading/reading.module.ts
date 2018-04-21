import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { Routes, RouterModule } from '@angular/router';
import { ReadingComponent } from './reading.component';
import { ClassToExcelReaderComponent } from './readers/class-to-excel-reader.component';
import { ClassToExcelRawReaderComponent } from './readers/class-to-excel-raw-reader.component';

const routes: Routes = [{
    path: "reading",
    children: [
      { path: "", component: ReadingComponent},
      { path: "ClassToExcelReader", component: ClassToExcelReaderComponent },
      { path: "ClassToExcelRawReader", component: ClassToExcelRawReaderComponent }
    ]
}];

@NgModule({
  imports: [
    RouterModule.forChild(routes),
    CommonModule
  ],
  declarations: [
    ReadingComponent,
    ClassToExcelReaderComponent,
    ClassToExcelRawReaderComponent
  ]
})
export class ReadingModule { }
