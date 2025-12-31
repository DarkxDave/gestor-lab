import { Routes } from '@angular/router';
import { FormTpaComponent } from './pages/form-tpa/form-tpa.component';
import { HomeComponent } from './pages/home/home.component';
import { FormRamComponent } from './pages/form-ram/form-ram.component';
import { FormRmylComponent } from './pages/form-rmyl/form-rmyl.component';
import { FormSalComponent } from './pages/form-sal/form-sal.component';
import { FormSaureusComponent } from './pages/form-saureus/form-saureus.component';
import { FormEnteroComponent } from './pages/form-entero/form-entero.component';
import { FormCtcfeComponent } from './pages/form-ctcfe/form-ctcfe.component';
import { SamplesComponent } from './pages/samples/samples.component';

export const routes: Routes = [
  {
    path: '',
    component: HomeComponent
  },
  {
    path: 'home',
    component: HomeComponent
  },
  {
    path: 'form-tpa',
    component: FormTpaComponent
  },
  {
    path: 'form-ram',
    component: FormRamComponent
  },
  {
    path: 'form-rmyl',
    component: FormRmylComponent
  },
  {
    path: 'form-sal',
    component: FormSalComponent
  },
  {
    path: 'form-saureus',
    component: FormSaureusComponent
  },
  {
    path: 'form-entero',
    component: FormEnteroComponent
  },
  {
    path: 'form-ctcfe',
    component: FormCtcfeComponent
  },
  {
    path: 'samples',
    component: SamplesComponent
  },
];
