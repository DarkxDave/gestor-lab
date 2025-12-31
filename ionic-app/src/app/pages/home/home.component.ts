import { Component, OnInit } from '@angular/core';
import { IonicModule, ActionSheetController } from '@ionic/angular';
import { FormsModule } from '@angular/forms';
import { CommonModule } from '@angular/common';
import { Router, RouterModule } from '@angular/router';

@Component({
  selector: 'app-home',
  templateUrl: './home.component.html',
  styleUrls: ['./home.component.scss'],
  standalone: true,
  imports: [IonicModule, FormsModule, CommonModule, RouterModule]
})
export class HomeComponent  implements OnInit {

  constructor(private router: Router, private actionSheetCtrl: ActionSheetController) { }

  ngOnInit() {}

  goToTpa() {
    this.router.navigate(['/form-tpa']);
  }

  async openSampleChoice() {
    const sheet = await this.actionSheetCtrl.create({
      header: '¿Qué deseas hacer?',
      buttons: [
        {
          text: 'Generar nueva muestra',
          icon: 'add-circle',
          handler: () => this.router.navigate(['/form-tpa'])
        },
        {
          text: 'Buscar muestra',
          icon: 'search',
          handler: () => this.router.navigate(['/samples'])
        },
        {
          text: 'Cancelar',
          role: 'cancel'
        }
      ]
    });
    await sheet.present();
  }

}
