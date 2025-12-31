import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { forkJoin, Observable } from 'rxjs';

@Injectable({
  providedIn: 'root'
})
export class FormService {

  private forms = ['tpa', 'ram', 'rmyl', 'sal', 'saureus', 'entero', 'ctcfe'];

  constructor(private http: HttpClient) { }

  saveAll(sampleId: string, formData: { [key: string]: any }): Observable<any[]> {
    const saveObservables = this.forms.map(formName => {
      const endpoint = `/api/form-${formName}/save`;
      const data = { ...formData[formName], sample_id: sampleId };
      return this.http.post(endpoint, data);
    });
    return forkJoin(saveObservables);
  }

  exportAll(sampleId: string): void {
    window.open(`/api/export/all-forms?sample_id=${sampleId}`, '_blank');
  }
}
