import { BrowserModule } from '@angular/platform-browser';
import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { HttpClientModule } from '@angular/common/http';

import { AppRoutingModule } from './app-routing.module';

import { AppComponent } from './app.component';
import { WelcomeComponent } from './welcome/welcome.component';
import { ReadingModule } from './reading/reading.module';
import { WritingModule } from './writing/writing.module';


@NgModule({
  declarations: [
    AppComponent,
    WelcomeComponent
  ],
  imports: [
    BrowserModule,    
    ReadingModule, // must be BEFORE AppRoutingModule to register child routes
    WritingModule,  // must be BEFORE AppRoutingModule to register child routes
    AppRoutingModule,
    CommonModule,
    HttpClientModule 
  ],
  providers: [],
  bootstrap: [AppComponent]
})
export class AppModule { }
