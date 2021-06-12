import { Component, OnInit } from '@angular/core';

import { AuthService } from '../auth.service';
import { GraphService } from '../graph.service';
import { AlertsService } from '../alerts.service';
import { NewEvent } from './new-event';
import { IExcelWorkSheetObj } from 'src/app/utility/app-excel-obj';
import { ExportExcelService } from 'src/app/shared/export-excel.service';
import { FormBuilder, FormControl, FormGroup, Validators } from '@angular/forms';

@Component({
  selector: 'app-new-event',
  templateUrl: './new-event.component.html',
  styleUrls: ['./new-event.component.scss']
})
export class NewEventComponent implements OnInit {

  model = new NewEvent();
  subjectInput = new FormControl();
  startInput = new FormControl();
  attendeesInput = new FormControl();
  endInput = new FormControl();
  locationInput = new FormControl();
  candidateNameInput = new FormControl();
  jobTitleInput = new FormControl();
  interviewerInput = new FormControl();
  scheduledByInput = new FormControl();

  mainStorage =[];
  storageValues = {
    SubjectInput: '',
    AttendeesInput: '',
    StartInput: '',
    EndInput: '',
    LocationInput: '',
    CandidateNameInput: '',
    JobTitleInput: '',
    InterviewerInput: '',
    ScheduledByInput: '',
  }

  constructor(
    private authService: AuthService,
    private graphService: GraphService,
    private exportExcelService: ExportExcelService,
    private alertsService: AlertsService) { }

  ngOnInit(): void {
    this.getValueChanges();
  }

  onSubmit(): void {
    const timeZone = this.authService.user?.timeZone ?? 'UTC';
    const graphEvent = this.model.getGraphEvent(timeZone);

    this.graphService.addEventToCalendar(graphEvent)
      .then(() => {
        this.alertsService.addSuccess('Event created.');
      }).catch(error => {
        this.alertsService.addError('Error creating event.', error.message);
      });

      this.mainStorage.push(this.storageValues);
      this.saveToLocalStorage();
  }

  getValueChanges(){
    this.subjectInput.valueChanges.subscribe((subjectInputValue) =>{
      this.storageValues.SubjectInput = subjectInputValue;
      JSON.stringify(this.storageValues);
    });
    this.attendeesInput.valueChanges.subscribe(attendeesInputValue =>{
      this.storageValues.AttendeesInput = attendeesInputValue;
      JSON.stringify(this.storageValues);
    });
    this.startInput.valueChanges.subscribe(startInputValue =>{
      this.storageValues.StartInput = startInputValue;
      JSON.stringify(this.storageValues);
    });
    this.endInput.valueChanges.subscribe(endInputValue =>{
      this.storageValues.EndInput = endInputValue;
      JSON.stringify(this.storageValues);
    });
    this.locationInput.valueChanges.subscribe(locationInputValue =>{
      this.storageValues.LocationInput = locationInputValue;
      JSON.stringify(this.storageValues);
    });
    this.candidateNameInput.valueChanges.subscribe(candidateNameInputValue =>{
      this.storageValues.CandidateNameInput = candidateNameInputValue;
      JSON.stringify(this.storageValues);
    });
    this.jobTitleInput.valueChanges.subscribe(jobTitleInputValue =>{
      this.storageValues.JobTitleInput = jobTitleInputValue;
      JSON.stringify(this.storageValues);
    });
    this.interviewerInput.valueChanges.subscribe(interviewerInputValue =>{
      this.storageValues.InterviewerInput = interviewerInputValue;
      JSON.stringify(this.storageValues);
    });
    this.scheduledByInput.valueChanges.subscribe(scheduledByInputValue =>{
      this.storageValues.ScheduledByInput = scheduledByInputValue;
      JSON.stringify(this.storageValues);
    });
  }

  saveToLocalStorage(){
    localStorage.setItem('interviewListData', JSON.stringify(this.mainStorage));
  }

  downloadAsExcel(){
    const mainArray = [];
    let interviewDataArray = [];
    interviewDataArray = JSON.parse(localStorage.getItem('interviewListData'));

    interviewDataArray.forEach(element =>{
      mainArray.push({
        Subject: element.SubjectInput,
        Attendees: element.AttendeesInput,
        'Candidate Name': element.CandidateNameInput,
        'Start Date': element.StartInput,
        'End Date': element.EndInput,
        'Job Title': element.JobTitleInput,
        Location: element.LocationInput,
        Interviewer: element.InterviewerInput,
        'Scheduled By': element.ScheduledByInput
      });
    });

    const excelWorkSheetObj: IExcelWorkSheetObj[] = [];
                excelWorkSheetObj.push({
                    WorkSheet_Obj: mainArray,
                    WorkSheet_Name: 'InterviewList'
                });
    this.exportExcelService.exportAsExcelWorksheets(excelWorkSheetObj, 'InterviewList');
  }
}