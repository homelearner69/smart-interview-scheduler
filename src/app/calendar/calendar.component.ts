import { Component, OnInit } from '@angular/core';
import * as moment from 'moment-timezone';
import { findIana } from 'windows-iana';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import { AuthService } from '../auth.service';
import { GraphService } from '../graph.service';
import { AlertsService } from '../alerts.service';
import { CalendarOptions } from '@fullcalendar/angular';

//import {Event, DateTimeTimeZone } from '../event';

@Component({
  selector: 'app-calendar',
  templateUrl: './calendar.component.html',
  styleUrls: ['./calendar.component.scss']
})
export class CalendarComponent implements OnInit {

  public events?: MicrosoftGraph.Event[];
  
  eventAry: any =[];

  calendarOptions: CalendarOptions ={};

  constructor(
    private authService: AuthService,
    private graphService: GraphService,
    private alertsService: AlertsService) { }

  ngOnInit() {
    // // Convert the user's timezone to IANA format
    // const ianaName = findIana(this.authService.user?.timeZone ?? 'UTC');
    // const timeZone = ianaName![0].valueOf() || this.authService.user?.timeZone || 'UTC';

    // // Get midnight on the start of the current week in the user's timezone,
    // // but in UTC. For example, for Pacific Standard Time, the time value would be
    // // 07:00:00Z
    // var startOfWeek = moment.tz(timeZone).startOf('week').utc();
    // var endOfWeek = moment(startOfWeek).add(7, 'day');

    // console.log(this.authService.user?.displayName);

    // this.graphService.getEvents()
    //   .then((events) => {
    //     this.events = events;
    //     console.log(events);
    this.graphService.getEvents()
      .then((events) => {
        this.events = events;
        console.log(events);

        this.events?.forEach(obj => {
          var title = obj.subject;
          //var date = obj.start?.dateTime?.split('T')[0];
          var start = obj.start?.dateTime;
          var end = obj.end?.dateTime;

          const ev = { title : title, start  : start, end : end};
          this.eventAry.push(ev);
        });
        this.calendarOptions;
        console.log(this.eventAry);
        
        this.calendarOptions = {
          initialView: 'dayGridMonth',
          events: this.eventAry
        }
      }).catch((error) => {
        console.log(error);
      });


    // this.graphService.getCalendarView(
    //   startOfWeek.format(),
    //   endOfWeek.format(),
    //   this.authService.user?.timeZone ?? 'UTC')
    //     .then((events) => {
    //       console.log(events);
    //       this.events = events;
    //       // Temporary to display raw results
    //       this.alertsService.addSuccess('Events from Graph', JSON.stringify(events, null, 2));
    //     }).catch(e => {
    //       console.log(e);
    //     });
  }

  formatDateTimeTimeZone(dateTime: MicrosoftGraph.DateTimeTimeZone | undefined | null): string {
    if (dateTime == undefined || dateTime == null) {
      return '';
    }
  
    try {
      // Pass UTC for the time zone because the value
      // is already adjusted to the user's time zone
      return moment.tz(dateTime.dateTime, 'UTC').format();
    }
    catch(error) {
      this.alertsService.addError('DateTimeTimeZone conversion error', JSON.stringify(error));
      return '';
    }
  }
}