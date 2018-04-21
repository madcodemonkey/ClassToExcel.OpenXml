import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { WritingComponent } from './writing.component';
import { Routes, RouterModule } from '@angular/router';

const routes : Routes = [{
  path: "writing",
  children: [
    { path: "", component: WritingComponent }
  ]
}];

@NgModule({
  imports: [
    RouterModule.forChild(routes),
    CommonModule
  ],
  declarations: [WritingComponent]
})
export class WritingModule { }
