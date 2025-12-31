import { Component, OnInit, ChangeDetectorRef, HostListener, ViewChild } from '@angular/core';
import { HttpClient, HttpClientModule } from '@angular/common/http';
import { ActivatedRoute, Router, RouterModule } from '@angular/router';
import { IonicModule, AlertController, IonContent } from '@ionic/angular';
import { FormsModule } from '@angular/forms';
import { CommonModule } from '@angular/common';
import { timeout, catchError } from 'rxjs/operators';
import { throwError } from 'rxjs';
import { FormTabsComponent } from '../../components/form-tabs/form-tabs.component';

@Component({
  selector: 'app-form-tpa',
  templateUrl: './form-tpa.component.html',
  styleUrls: ['./form-tpa.component.scss'],
  standalone: true,
  imports: [IonicModule, FormsModule, CommonModule, RouterModule, HttpClientModule, FormTabsComponent]
})
export class FormTpaComponent implements OnInit {
  @ViewChild(IonContent) ionContent?: IonContent;
  sampleId: string = '';
  formData: any = {};
  message: string = '';
  isEditMode: boolean = false;
  isLoading: boolean = false;
  private originalFormData: string = '';

  constructor(
    private http: HttpClient,
    private route: ActivatedRoute,
    private cdr: ChangeDetectorRef,
    public router: Router,
    private alertController: AlertController
  ) {}

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
      this.formData = { sample_id: '' };
      this.originalFormData = JSON.stringify(this.formData);
    }
    // Aplicar clases has-value después de cargar
    setTimeout(() => this.updateInputClasses(), 500);
  }

  @HostListener('window:beforeunload', ['$event'])
  unloadNotification($event: any): void {
    if (this.hasChanges()) {
      $event.returnValue = true;
    }
  }

  private hasChanges(): boolean {
    const currentData = JSON.stringify(this.formData);
    return currentData !== this.originalFormData;
  }

  private updateInputClasses(): void {
    // Aplicar clase has-value a inputs con contenido
    const inputs = document.querySelectorAll('ion-input, ion-textarea');
    inputs.forEach((input: Element) => {
      const ionInput = input as any;
      if (ionInput.value || ionInput.getAttribute('value')) {
        ionInput.classList.add('has-value');
      } else {
        ionInput.classList.remove('has-value');
      }
    });
  }

  onInputChange(): void {
    // Se ejecuta cuando cambia cualquier input
    setTimeout(() => this.updateInputClasses(), 50);
  }

  private scrollToTop(): void {
    if (this.ionContent) {
      this.ionContent.scrollToTop(400);
    } else if (typeof window !== 'undefined' && window.scrollTo) {
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }
  }

  async canDeactivate(): Promise<boolean> {
    if (!this.hasChanges()) {
      return true;
    }

    const alert = await this.alertController.create({
      header: 'Cambios sin guardar',
      message: 'Se han generado cambios. ¿Seguro que desea salir sin guardar?',
      buttons: [
        {
          text: 'Guardar',
          role: 'save',
          handler: () => {
            this.saveForm();
            return false;
          }
        },
        {
          text: 'Salir sin guardar',
          role: 'confirm',
          handler: () => {
            this.router.navigate(['/']);
            return true;
          }
        }
      ]
    });

    await alert.present();
    const { role } = await alert.onDidDismiss();
    return role === 'confirm';
  }

  async onGoHome(): Promise<void> {
    const canLeave = await this.canDeactivate();
    if (canLeave) {
      await this.router.navigate(['/']);
    }
  }

  async onBack(event?: Event): Promise<void> {
    if (event) {
      event.preventDefault();
    }
    const canLeave = await this.canDeactivate();
    if (canLeave) {
      await this.router.navigate(['/']);
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
    console.log('Iniciando carga de sample_id:', idToLoad);
    
    this.http.get<any>(`/api/form-tpa?sample_id=${idToLoad}`, { 
      responseType: 'json' 
    }).pipe(
      timeout(8000), // Timeout de 8 segundos
      catchError(err => {
        if (err.name === 'TimeoutError') {
          this.message = '⏱️ La carga está tardando demasiado. Por favor, verifica tu conexión a la base de datos.';
          this.isLoading = false;
          this.cdr.markForCheck();
        }
        return throwError(() => err);
      })
    ).subscribe({
      next: (data) => {
        console.log('Datos recibidos:', data);
        this.isLoading = false;
        if (data && data.data) {
          this.formData = { ...data.data };
          this.originalFormData = JSON.stringify(this.formData);
          this.message = '✅ Datos cargados correctamente';
          if (!this.isEditMode) {
            this.isEditMode = true;
            this.sampleId = idToLoad;
          }
        } else if (data && data.data === null) {
          this.message = '⚠️ Sample ID no encontrado';
        } else {
          this.message = '⚠️ No se encontraron datos para este sample_id';
        }
        this.cdr.markForCheck();
      },
      error: (error) => {
        console.error('Error al cargar:', error);
        this.isLoading = false;
        if (error.status === 504) {
          this.message = '⏱️ Timeout: La base de datos tardó demasiado en responder';
        } else if (error.status === 404) {
          this.message = '❌ Sample ID no encontrado en la base de datos';
        } else {
          this.message = '❌ Error al cargar: ' + (error?.error?.message || error?.message || 'Error desconocido');
        }
        this.cdr.markForCheck();
      }
    });
  }

  saveForm() {
    const idToSave = this.isEditMode ? this.sampleId : this.formData.sample_id;
    if (!idToSave) {
      this.message = '⚠️ Por favor, ingrese un ID de muestra.';
      return;
    }

    this.isLoading = true;
    this.message = '💾 Guardando...';
    this.cdr.markForCheck();

    this.http.post<any>('/api/form-tpa/save', { ...this.formData, sample_id: idToSave }).pipe(
      timeout(10000),
      catchError(err => {
        if (err.name === 'TimeoutError') {
          this.message = '⏱️ El guardado está tardando demasiado. Verifica tu conexión.';
          this.isLoading = false;
          this.cdr.markForCheck();
        }
        return throwError(() => err);
      })
    ).subscribe({
      next: (response) => {
        this.isLoading = false;
        this.message = '✅ ' + response.message;
        this.originalFormData = JSON.stringify(this.formData);
        if (!this.isEditMode) {
          this.isEditMode = true;
          this.sampleId = idToSave;
        }
        this.scrollToTop();
        this.cdr.markForCheck();
      },
      error: (error) => {
        this.isLoading = false;
        this.message = '❌ Error al guardar: ' + (error?.error?.message || error?.message || 'Error desconocido');
        console.error(error);
        this.cdr.markForCheck();
      }
    });
  }

  exportToExcel() {
    const idToExport = this.isEditMode ? this.sampleId : this.formData.sample_id;
    if (!idToExport) {
      this.message = 'No hay un ID de muestra para exportar.';
      return;
    }
    
    window.location.href = `/api/export/all?sample_id=${idToExport}`;
  }
}
