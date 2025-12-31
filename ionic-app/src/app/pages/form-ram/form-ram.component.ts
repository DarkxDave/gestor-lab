import { Component, OnInit, AfterViewInit, ChangeDetectorRef, ViewChild } from '@angular/core';
import { HttpClient, HttpClientModule } from '@angular/common/http';
import { ActivatedRoute, RouterModule } from '@angular/router';
import { IonicModule, IonContent } from '@ionic/angular';
import { FormsModule } from '@angular/forms';
import { CommonModule } from '@angular/common';
import { timeout, catchError } from 'rxjs/operators';
import { throwError } from 'rxjs';
import { FormTabsComponent } from '../../components/form-tabs/form-tabs.component';

@Component({
  selector: 'app-form-ram',
  templateUrl: './form-ram.component.html',
  styleUrls: ['./form-ram.component.scss'],
  standalone: true,
  imports: [IonicModule, FormsModule, CommonModule, RouterModule, HttpClientModule, FormTabsComponent]
})
export class FormRamComponent implements OnInit, AfterViewInit {
  @ViewChild(IonContent) ionContent?: IonContent;
  sampleId: string = '';
  formData: any = {};
  message: string = '';
  isEditMode: boolean = false;
  isLoading: boolean = false;
  
  // Calculated results
  resultCFU: string = '—';
  dilutionAcceptance: Array<{dilution: number, status: string}> = [];

  constructor(private http: HttpClient, private route: ActivatedRoute, private cdr: ChangeDetectorRef) {}

  get currentSampleId(): string {
    return this.isEditMode ? this.sampleId : this.formData.sample_id || '';
  }

  set currentSampleId(value: string) {
    if (this.isEditMode) {
      this.sampleId = value;
    } else {
      this.formData.sample_id = value;
    }
  }

  ngOnInit() {
    const id = this.route.snapshot.queryParamMap.get('sample_id');
    if (id) {
      this.isEditMode = true;
      this.sampleId = id;
      this.loadSampleData();
    } else {
      this.isEditMode = false;
      this.initializeFormData();
    }
  }

  ngAfterViewInit() {
    // Recalculate when component view is ready
    setTimeout(() => this.recalculate(), 100);
  }

  initializeFormData() {
    this.formData = {
      sample_id: '',
      inicio_incubacion_fecha: '',
      inicio_incubacion_hora: '',
      inicio_incubacion_analista: '',
      termino_analisis_fecha: '',
      termino_analisis_hora: '',
      termino_analisis_analista: '',
      cc2_pesado_temp: '',
      cc2_pesado_ufc: '',
      cc2_siembra: '',
      cc2_hora_inicio: '',
      cc2_hora_termino: '',
      cc2_temp: '',
      cc2_ecoli_ufc: '',
      cc2_blanco_ufc: '',
      siembra_tiempo_ok: false,
      siembra_n_muestra_10g_90ml: '',
      siembra_n_muestra_50g_450ml: '',
      cc_duplicado_ali_analisis: '',
      cc_duplicado_ali_cumple: '',
      cc_control_pos_blanco_ali_analisis: '',
      cc_control_pos_blanco_ali_cumple: '',
      cc_control_siembra_ali_analisis: '',
      cc_control_siembra_ali_cumple: '',
      mic_desfavorable_si: false,
      mic_desfavorable_no: false,
      mic_tabla_pagina: '',
      mic_limite: '',
      mic_fecha_entrega: '',
      mic_hora_entrega: '',
      datos_suspension_inicial_den: '',
      datos_volumen_petri_ml: '',
      datos_dilucion_log10_1: '',
      observaciones: ''
    };
    
    // Initialize dilution rows (5 rows)
    for (let i = 1; i <= 5; i++) {
      this.formData[`datos_colonias_num_a_${i}`] = '';
      this.formData[`datos_colonias_num_b_${i}`] = '';
      this.formData[`datos_colonias_por_conf_a_${i}`] = '';
      this.formData[`datos_colonias_por_conf_b_${i}`] = '';
      this.formData[`datos_colonias_conf_a_${i}`] = '';
      this.formData[`datos_colonias_conf_b_${i}`] = '';
      this.formData[`datos_colonias_final_a_${i}`] = '';
      this.formData[`datos_colonias_final_b_${i}`] = '';
    }
  }

  private scrollToTop(): void {
    if (this.ionContent) {
      this.ionContent.scrollToTop(400);
    } else if (typeof window !== 'undefined' && window.scrollTo) {
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }
  }

  loadSampleData() {
    const idToLoad = this.isEditMode ? this.sampleId : this.formData.sample_id;
    if (!idToLoad) {
      this.message = 'Por favor, ingrese un ID de muestra.';
      return;
    }
    
    this.isLoading = true;
    this.message = 'Cargando...';
    this.cdr.markForCheck();
    
    this.http.get<any>(`/api/form-ram?sample_id=${idToLoad}`).pipe(
      timeout(8000),
      catchError(err => {
        if (err.name === 'TimeoutError') {
          this.message = '⏱️ La carga está tardando demasiado.';
          this.isLoading = false;
          this.cdr.markForCheck();
        }
        return throwError(() => err);
      })
    ).subscribe({
      next: (data) => {
        this.isLoading = false;
        if (data && data.data) {
          this.formData = this.formatLoadedData(data.data);
          this.message = '✅ Datos cargados correctamente';
          if (!this.isEditMode) {
            this.isEditMode = true;
            this.sampleId = idToLoad;
          }
          setTimeout(() => this.recalculate(), 100);
        } else {
          this.message = 'ℹ️ No hay datos previos para este ALI; complete el formulario y guarde.';
          this.initializeFormData();
          this.formData.sample_id = idToLoad;
        }
        this.cdr.markForCheck();
      },
      error: (err) => {
        this.isLoading = false;
        console.error('Error loading data:', err);
        if (err?.status === 404) {
          // Primera vez: no existe registro, habilitar captura sin mostrar error
          this.initializeFormData();
          this.formData.sample_id = idToLoad;
          this.message = '';
          this.cdr.markForCheck();
          return;
        }
        this.message = '❌ Error al cargar los datos';
        this.cdr.markForCheck();
      }
    });
  }

  formatLoadedData(data: any): any {
    const formatted = { ...data };
    
    // Format dates
    const dateFields = ['inicio_incubacion_fecha', 'termino_analisis_fecha', 'mic_fecha_entrega'];
    dateFields.forEach(field => {
      if (formatted[field]) {
        formatted[field] = this.formatDate(formatted[field]);
      }
    });
    
    // Format times
    const timeFields = ['inicio_incubacion_hora', 'termino_analisis_hora', 'cc2_hora_inicio', 'cc2_hora_termino', 'mic_hora_entrega'];
    timeFields.forEach(field => {
      if (formatted[field]) {
        formatted[field] = this.formatTime(formatted[field]);
      }
    });
    
    return formatted;
  }

  formatDate(value: any): string {
    if (!value) return '';
    try {
      if (value instanceof Date) {
        const tz = value.getTimezoneOffset();
        const local = new Date(value.getTime() - tz * 60000);
        return local.toISOString().slice(0, 10);
      }
      const s = String(value);
      const m = s.match(/^\d{4}-\d{2}-\d{2}/);
      if (m) return m[0];
      if (s.includes('T')) return s.slice(0, 10);
      if (s.includes(' ')) return s.split(' ')[0];
      return s.slice(0, 10);
    } catch (_) {
      return '';
    }
  }

  formatTime(value: any): string {
    if (!value) return '';
    if (value instanceof Date) {
      const hh = String(value.getHours()).padStart(2, '0');
      const mm = String(value.getMinutes()).padStart(2, '0');
      return `${hh}:${mm}`;
    }
    const s = String(value);
    if (/^\d{2}:\d{2}/.test(s)) return s.slice(0, 5);
    return '';
  }

  saveForm() {
    if (!this.sampleId && !this.formData.sample_id) {
      this.message = '⚠️ Por favor, ingrese un ID de muestra';
      return;
    }
    
    this.isLoading = true;
    this.message = '💾 Guardando...';
    this.cdr.markForCheck();
    const dataToSave = { ...this.formData, sample_id: this.sampleId || this.formData.sample_id };
    
    this.http.post<any>('/api/form-ram/save', dataToSave).pipe(
      timeout(10000),
      catchError(err => {
        if (err.name === 'TimeoutError') {
          this.message = '⏱️ El guardado está tardando demasiado.';
          this.isLoading = false;
          this.cdr.markForCheck();
        }
        return throwError(() => err);
      })
    ).subscribe({
      next: (response) => {
        this.isLoading = false;
        this.message = '✅ ' + (response.message || 'Guardado correctamente');
        this.scrollToTop();
        this.cdr.markForCheck();
      },
      error: (err) => {
        this.isLoading = false;
        console.error('Error saving:', err);
        this.message = '❌ Error al guardar';
        this.cdr.markForCheck();
      }
    });
  }

  exportToExcel() {
    const id = this.sampleId || this.formData.sample_id;
    if (!id) {
      this.message = '⚠️ Debe cargar una muestra primero';
      return;
    }
    window.location.href = `/api/export/all?sample_id=${encodeURIComponent(id)}`;
  }

  onInputChange() {
    this.recalculate();
  }

  // Calculation methods from original EJS
  private parseNum(v: any): number | null {
    if (v == null) return null;
    const s = String(v).trim().replace(',', '.');
    if (s === '') return null;
    const n = Number(s);
    return Number.isFinite(n) ? n : null;
  }

  private chiCDF_k1(x: number): number {
    if (x <= 0 || !Number.isFinite(x)) return 0;
    const t = Math.sqrt(x / 2);
    
    const erf = (z: number): number => {
      const sign = z < 0 ? -1 : 1;
      const a1 = 0.254829592, a2 = -0.284496736, a3 = 1.421413741, a4 = -1.453152027, a5 = 1.061405429, p = 0.3275911;
      const x = Math.abs(z);
      const t = 1 / (1 + p * x);
      const y = 1 - ((((a5 * t + a4) * t + a3) * t + a2) * t + a1) * t * Math.exp(-x * x);
      return sign * y;
    };
    
    return Math.max(0, Math.min(1, erf(t)));
  }

  private countTA(vals: any[]): number {
    return vals.filter(v => v !== null && v !== '').length;
  }

  recalculate() {
    const vol = this.parseNum(this.formData.datos_volumen_petri_ml);
    const den = this.parseNum(this.formData.datos_suspension_inicial_den);
    
    // Calculate final colonies for each dilution row
    const C: any[] = [null], D: any[] = [null], F: any[] = [null], G: any[] = [null];
    const H: any[] = [null], I: any[] = [null], J: any[] = [null], K: any[] = [null];
    
    for (let i = 1; i <= 5; i++) {
      C[i] = this.parseNum(this.formData[`datos_colonias_num_a_${i}`]);
      D[i] = this.parseNum(this.formData[`datos_colonias_num_b_${i}`]);
      F[i] = this.parseNum(this.formData[`datos_colonias_por_conf_a_${i}`]);
      G[i] = this.parseNum(this.formData[`datos_colonias_por_conf_b_${i}`]);
      H[i] = this.parseNum(this.formData[`datos_colonias_conf_a_${i}`]);
      I[i] = this.parseNum(this.formData[`datos_colonias_conf_b_${i}`]);
      
      // Calculate J (A - final)
      let j: any = '';
      if (F[i] != null && C[i] != null && F[i] > C[i]) j = 'ERROR';
      else if ((F[i] == null) !== (H[i] == null)) j = 'ERROR';
      else if (F[i] == null && H[i] == null) j = '';
      else if (F[i] != null && H[i] != null) {
        if (F[i] < H[i]) j = 'ERROR';
        else j = C[i] != null ? C[i] * (H[i] / (F[i] === 0 ? 1 : F[i])) : '';
      }
      J[i] = j;
      this.formData[`datos_colonias_final_a_${i}`] = (j === 'ERROR' || j === '') ? j : String(Math.round(j * 1000) / 1000);
      
      // Calculate K (B - final)
      let k: any = '';
      if (G[i] != null && D[i] != null && G[i] > D[i]) k = 'ERROR';
      else if ((G[i] == null) !== (I[i] == null)) k = 'ERROR';
      else if (G[i] == null && I[i] == null) k = '';
      else if (G[i] != null && I[i] != null) {
        if (G[i] < I[i]) k = 'ERROR';
        else k = D[i] != null ? D[i] * (I[i] / (G[i] === 0 ? 1 : G[i])) : '';
      }
      K[i] = k;
      this.formData[`datos_colonias_final_b_${i}`] = (k === 'ERROR' || k === '') ? k : String(Math.round(k * 1000) / 1000);
    }
    
    // Validation checks
    const q25 = (vol != null && den != null);
    const q26 = !!((C[1] != null || D[1] != null || F[1] != null || G[1] != null) ||
                   (C[2] != null || D[2] != null || F[2] != null || G[2] != null));
    const anyFinal12 = !!((J[1] !== '' && J[1] !== 'ERROR') || (K[1] !== '' && K[1] !== 'ERROR') ||
                          (J[2] !== '' && J[2] !== 'ERROR') || (K[2] !== '' && K[2] !== 'ERROR'));
    
    // Count plates per dilution
    const countPlates = (i: number) => anyFinal12 ? this.countTA([F[i], G[i]]) : this.countTA([C[i], D[i]]);
    const q40 = countPlates(1);
    const q41 = countPlates(2);
    const q45 = countPlates(3);
    const q46 = countPlates(4);
    const q47 = countPlates(5);
    
    // Sum final colonies
    const hasFinalAny = [1, 2, 3, 4, 5].some(i => (J[i] !== '' && J[i] !== 'ERROR') || (K[i] !== '' && K[i] !== 'ERROR'));
    const sumFinal = hasFinalAny
      ? [1, 2, 3, 4, 5].reduce((s, i) => s + (Number(J[i]) || 0) + (Number(K[i]) || 0), 0)
      : [1, 2, 3, 4, 5].reduce((s, i) => s + (Number(C[i]) || 0) + (Number(D[i]) || 0), 0);
    
    // Q39 dilution factor
    let q39 = null;
    if (den != null) {
      q39 = (den === 0) ? null : (1 / den);
    }
    
    // Q42 weighted volume
    const q42 = (vol != null) ? vol * (q40 + 0.1 * q41 + 0.01 * q45 + 0.001 * q46 + 0.0001 * q47) : null;
    
    // Calculate result
    let res = null;
    if (q25 && q26 && q39 != null && q42 != null) {
      if (sumFinal === 0) res = 1 / (q39 * (vol || 1));
      else res = sumFinal / (q42 * q39);
    }
    
    this.resultCFU = (res == null || !Number.isFinite(res))
      ? '—'
      : Number(res).toExponential(1).replace('e+', 'E+').replace('e-', 'E-');
    
    // Dilution acceptance (chi-square test)
    const alpha = 0.01;
    this.dilutionAcceptance = [];
    
    const perRowOK = (i: number): string => {
      const a = C[i], b = D[i];
      const n = this.countTA([a, b]);
      if (res == null || !Number.isFinite(res)) return 'NOT APPLICABLE';
      if ((a == null && b == null) || n < 2) return 'NOT APPLICABLE';
      const min = Math.min(a || 0, b || 0), max = Math.max(a || 0, b || 0);
      if (max === 0) return 'NOT APPLICABLE';
      const avg = (min + max) / 2;
      const x = 2 * (min * Math.log((min === 0 ? 1 : min) / avg) + max * Math.log(max / avg));
      const p = 1 - this.chiCDF_k1(x);
      return (p > alpha) ? 'YES' : 'NO';
    };
    
    for (let i = 1; i <= 5; i++) {
      this.dilutionAcceptance.push({
        dilution: i,
        status: perRowOK(i)
      });
    }
  }
}
