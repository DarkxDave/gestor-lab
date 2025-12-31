import { Component } from '@angular/core';
import { Router, NavigationStart, NavigationEnd, NavigationCancel, NavigationError } from '@angular/router';
import { CommonModule } from '@angular/common';
import { IonApp, IonRouterOutlet } from '@ionic/angular/standalone';
import { HttpClientModule } from '@angular/common/http';
import { FormsModule } from '@angular/forms';

@Component({
  selector: 'app-root',
  templateUrl: 'app.component.html',
  standalone: true,
  imports: [IonApp, IonRouterOutlet, CommonModule, HttpClientModule, FormsModule],
})
export class AppComponent {
  constructor(private router: Router) {
    this.router.events.subscribe(evt => {
      if (evt instanceof NavigationStart) {
        console.log('[Router] Start ->', evt.url);
      } else if (evt instanceof NavigationEnd) {
        console.log('[Router] End   ->', evt.url);
      } else if (evt instanceof NavigationCancel) {
        console.warn('[Router] Cancel ->', evt.url, evt.reason);
      } else if (evt instanceof NavigationError) {
        console.error('[Router] Error  ->', evt.error);
      }
    });
  }
}
