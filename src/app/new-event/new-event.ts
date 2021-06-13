import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { FormBuilder, FormControl, FormGroup, Validators } from '@angular/forms';
import { DatePipe } from '@angular/common';

// Model for the new event form
export class NewEvent {
  subject?: string;
  attendees?: string;
  location?: string;
  candidateName?: string;
  jobTitle?: string;
  interviewer?: string;
  scheduledBy?: string;
  start?: string;
  end?: string;
  body?: string;

  subjectInput = new FormControl();
  attendeesInput = new FormControl();
  endInput = new FormControl();
  startInput =  new FormControl();
  locationInput = new FormControl();
  candidateNameInput = new FormControl();
  jobTitleInput = new FormControl();
  interviewerInput = new FormControl();
  scheduledByInput = new FormControl();

  storageValues = {
    SubjectInput: '',
    AttendeesInput: '',
    StartInput: '',
  }

  constructor(
    private datePipe: DatePipe
  ) { }


  // Generate a MicrosoftGraph.Event from the model
  getGraphEvent(timeZone: string): MicrosoftGraph.Event {
    const graphEvent: MicrosoftGraph.Event = {
      subject: this.subject,
      start: {
        dateTime: this.start,
        timeZone: timeZone
      },
      end: {
        dateTime: this.end,
        timeZone: timeZone
      }
    };

    // If there are attendees, convert to array
    // and add them
    if (this.attendees && this.attendees.length > 0) {
      graphEvent.attendees = [];

      const emails = this.attendees.split(';');
      emails.forEach(email => {
        graphEvent.attendees?.push({
          type: 'required',
          emailAddress: {
            address: email
          }
        });
      });
    }
    

    // If there is a body, add it as plain text
    // if (this.body && this.body.length > 0) {
      graphEvent.body = {
        contentType: 'text',
        content: 'Dear ' + this.candidateName + '\n\nThank you for applying to Keysight' +'.'
        + '\n\nMy name is ' + this.scheduledBy + 'and I am the hiring manager at Keysight'
        + '\n\nI would like to schedule an online interview with you to discuss about your application for the ' + this.jobTitle + ' role'
        + '\n\nWould you be available for a interview from ' + this.datePipe.transform(this.start, 'MMM d,y, h:mmm:ss a') + ' to ' + this.datePipe.transform(this.end, 'MMM d,y, h:mmm:ss a')
        + '\n\nPlease let us know if you are available by accepting / request for another time.'
        + '\n\nLooking forward to hearing from you'
      };
    // }

    return graphEvent;
  }
}