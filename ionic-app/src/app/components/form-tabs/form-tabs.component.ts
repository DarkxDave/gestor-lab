import { Component, Input, OnInit, OnDestroy, ChangeDetectorRef } from '@angular/core';
import { Router, NavigationEnd, ActivatedRoute } from '@angular/router';
import { IonicModule } from '@ionic/angular';
import { FormsModule } from '@angular/forms';
import { CommonModule } from '@angular/common';
import { RouterModule } from '@angular/router';
import { Subject } from 'rxjs';
import { filter, takeUntil } from 'rxjs/operators';

@Component({
  selector: 'app-form-tabs',
  templateUrl: './form-tabs.component.html',
  styleUrls: ['./form-tabs.component.scss'],
  standalone: true,
  imports: [IonicModule, FormsModule, CommonModule, RouterModule]
})
export class FormTabsComponent implements OnInit, OnDestroy {
  @Input() sampleId: string = '';
  active: string = '';
  private destroy$ = new Subject<void>();
  private navigationInProgress = false;

  constructor(
    private router: Router,
    private route: ActivatedRoute,
    private cdr: ChangeDetectorRef
  ) { }

  ngOnInit() {
    // Detectar la ruta actual
    this.updateActiveTab();
    
    // Escuchar cambios de ruta para actualizar el segment dinámicamente
    this.router.events.pipe(
      filter(event => event instanceof NavigationEnd),
      takeUntil(this.destroy$)
    ).subscribe((event: any) => {
      console.log('NavigationEnd event detected, updating active tab');
      // Pequeño delay para asegurar que la ruta se actualizó
      setTimeout(() => {
        this.updateActiveTab();
        this.navigationInProgress = false;
        this.cdr.detectChanges();
      }, 100);
    });
  }

  private updateActiveTab() {
    const urlSegments = this.router.url.split('?')[0]; // Remover query params
    const formRoute = urlSegments.split('/')[1]; // "/form-tpa" -> "form-tpa"
    
    if (formRoute.startsWith('form-')) {
      const newActive = formRoute.replace('form-', ''); // "form-tpa" -> "tpa"
      if (newActive !== this.active) {
        this.active = newActive;
        console.log('Active tab updated to:', this.active);
        this.cdr.detectChanges();
      }
    }
  }

  ngOnDestroy() {
    this.destroy$.next();
    this.destroy$.complete();
  }

  segmentChanged(event: any) {
    if (this.navigationInProgress) {
      console.log('Navigation already in progress, ignoring segment change');
      return;
    }

    const form = event.detail.value;
    console.log('Tab changed to:', form, 'from:', this.active, 'with sampleId:', this.sampleId);
    
    if (form !== this.active) {
      if (!this.sampleId) {
        console.warn('No sampleId available for navigation');
        return;
      }

      this.navigationInProgress = true;
      console.log('Navigating to:', `/form-${form}`, 'with sample_id:', this.sampleId);
      
      this.router.navigate([`/form-${form}`], {
        queryParams: { sample_id: this.sampleId }
      }).then(success => {
        console.log('Navigation successful:', success);
      }).catch(error => {
        console.error('Navigation error:', error);
        this.navigationInProgress = false;
      });
    }
  }
}
