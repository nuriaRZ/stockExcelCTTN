import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { ExportExcelComponent } from './pages/export-excel/export-excel.component'

const routes: Routes = [
  { path: '', redirectTo: 'exportExcel', pathMatch: 'full'},  
  { path: 'exportExcel', component: ExportExcelComponent},
  { path: '**', component: ExportExcelComponent},

];

@NgModule({
  imports: [RouterModule.forRoot(routes)],
  exports: [RouterModule]
})
export class AppRoutingModule { }
