import { Component, OnInit } from '@angular/core';
import { HttpClient, HttpClientModule } from '@angular/common/http';
import { ActivatedRoute, RouterModule } from '@angular/router';
import { FormService } from '../../services/form';
import { IonicModule } from '@ionic/angular';
import { FormsModule } from '@angular/forms';
import { CommonModule } from '@angular/common';
import { FormTabsComponent } from 'src/app/components/form-tabs/form-tabs.component';

@Component({
  selector: 'app-form-rmyl',
  templateUrl: './form-rmyl.component.html',
  styleUrls: ['./form-rmyl.component.scss'],
  standalone: true,
  imports: [IonicModule, FormsModule, CommonModule, RouterModule, HttpClientModule, FormTabsComponent]
})
export class FormRmylComponent implements OnInit {
  sampleId: string = '';
  formData: any = {};
  message: string = '';
  allFormData: { [key: string]: any } = {};

  constructor(
    private http: HttpClient,
    private route: ActivatedRoute,
    private formService: FormService
  ) { }

  ngOnInit() {
    this.route.queryParams.subscribe(params => {
      if (params['sample_id']) {
        this.sampleId = params['sample_id'];
        this.loadSampleData();
      }
    });
  }

  loadSampleData() {
    if (this.sampleId) {
      this.http.get<any>(`/api/form-rmyl?sample_id=${this.sampleId}`).subscribe(
        data => {
          this.formData = data.data || {};
          this.allFormData['rmyl'] = this.formData;
          this.message = data.message || '';
        },
        error => {
          this.message = 'Error al cargar los datos';
          console.error(error);
        }
      );
    }
  }

  saveForm() {
    this.http.post<any>('/api/form-rmyl/save', { ...this.formData, sample_id: this.sampleId }).subscribe(
      response => {
        this.message = response.message;
      },
      error => {
        this.message = 'Error al guardar los datos';
        console.error(error);
      }
    );
  }
}
