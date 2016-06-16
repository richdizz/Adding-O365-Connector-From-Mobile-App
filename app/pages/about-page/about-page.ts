import {Component, NgZone} from '@angular/core';
import {NavController} from 'ionic-angular';

@Component({
  templateUrl: 'build/pages/about-page/about-page.html'
})
export class AboutPage {
  zone:NgZone;
  sub = {group: "undefined", webhook: "undefined"};

  constructor(private navController: NavController, zone: NgZone) {
    this.zone = zone;
  }
  

  connectToO365() {
    let ctrl = this;
    let ref = cordova.InAppBrowser.open("https://outlook.office.com/connectors/Connect?state=myAppsState&app_id=4b543361-dbc2-4726-a351-c4b43711d6c5&callback_url=https://localhost:44300/callback", "_blank", "location=no,clearcache=yes,toolbar=yes");
    ref.addEventListener("loadstart", function(evt) {
      ctrl.zone.run(() => {
        if (evt.url.indexOf('https://localhost:44300/callback') == 0) {
          //decode the uri into the response parameters
          let parts = evt.url.split('&');
          ctrl.sub.webhook = decodeURIComponent(parts[1].split('=')[1]);
          ctrl.sub.group = decodeURIComponent(parts[2].split('=')[1]);
          ref.close();
        }
      });
    });
  }
}
