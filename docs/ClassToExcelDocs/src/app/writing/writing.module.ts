import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { WritingComponent } from './writing.component';
import { Routes, RouterModule } from '@angular/router';
import { ExampleModule } from '../example/example.module';
import { ClassToExcelWriterComponent } from './writers/class-to-excel-writer.component';

const routes : Routes = [{
  path: "writing",
  children: [
    { path: "", component: WritingComponent },
    { path: "ClassToExcelWriter", component: ClassToExcelWriterComponent }

  ]
}];

@NgModule({
  imports: [
    RouterModule.forChild(routes),
    CommonModule,
    ExampleModule
  ],
  declarations: [WritingComponent, ClassToExcelWriterComponent]
})
export class WritingModule { }
