import { Component, OnInit, ChangeDetectorRef } from '@angular/core';
import { HttpClient, HttpClientModule } from '@angular/common/http';
import { IonicModule } from '@ionic/angular';
import { FormsModule } from '@angular/forms';
import { CommonModule } from '@angular/common';
import { RouterModule } from '@angular/router';
import { timeout, catchError } from 'rxjs/operators';
import { throwError } from 'rxjs';

@Component({
  selector: 'app-samples',
  templateUrl: './samples.component.html',
  styleUrls: ['./samples.component.scss'],
  standalone: true,
  imports: [IonicModule, FormsModule, CommonModule, RouterModule, HttpClientModule]
})
export class SamplesComponent implements OnInit {
  searchTerm: string = '';
  samples: any[] = [];
  isLoading: boolean = false;
  message: string = '';
  filteredSamples: any[] = [];

  constructor(private http: HttpClient, private cdr: ChangeDetectorRef) { }

  ngOnInit() {
    this.loadSamples();
  }

  loadSamples() {
    this.isLoading = true;
    this.message = 'Cargando muestras...';
    this.cdr.markForCheck();

    this.http.get<any>('/api/samples').pipe(
      timeout(8000),
      catchError(err => {
        if (err.name === 'TimeoutError') {
          this.message = '⏱️ La carga está tardando demasiado';
        } else {
          this.message = '❌ Error al cargar las muestras';
        }
        this.isLoading = false;
        this.cdr.markForCheck();
        return throwError(() => err);
      })
    ).subscribe({
      next: (response: any) => {
        this.isLoading = false;
        if (response && Array.isArray(response)) {
          this.samples = response;
          this.filteredSamples = response;
          if (this.samples.length === 0) {
            this.message = '⚠️ No hay muestras registradas';
          } else {
            this.message = `✅ ${this.samples.length} muestra(s) cargada(s)`;
          }
        } else if (response && response.data && Array.isArray(response.data)) {
          this.samples = response.data;
          this.filteredSamples = response.data;
          this.message = `✅ ${this.samples.length} muestra(s) cargada(s)`;
        } else {
          this.message = '⚠️ Formato de respuesta inesperado';
        }
        this.cdr.markForCheck();
      },
      error: (error) => {
        this.isLoading = false;
        console.error('Error loading samples:', error);
        this.message = '❌ Error al cargar las muestras: ' + (error?.message || 'Error desconocido');
        this.cdr.markForCheck();
      }
    });
  }

  searchSamples() {
    if (!this.searchTerm.trim()) {
      this.filteredSamples = this.samples;
      return;
    }
    
    const term = this.searchTerm.toLowerCase();
    this.filteredSamples = this.samples.filter(sample =>
      sample.sample_id.toLowerCase().includes(term)
    );

    // Ordenar de menor a mayor por ID
    this.filteredSamples.sort((a, b) => a.id - b.id);

    if (this.filteredSamples.length === 0) {
      this.message = `⚠️ No se encontraron muestras con: "${this.searchTerm}"`;
    } else {
      this.message = `✅ ${this.filteredSamples.length} resultado(s)`;
    }
  }

  clearSearch() {
    this.searchTerm = '';
    this.filteredSamples = this.samples;
    this.message = `✅ ${this.samples.length} muestra(s) cargada(s)`;
  }
}
