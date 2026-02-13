export interface IEvent {
  Id: number;
  Title: string;
  EventDate: string; // ISO string
  EndDate: string; // ISO string
  Description: string;
  Category: string;
  CategoryColor: string;
  TargetAudience: string;
  Location: string;
  Department: string;
  Contact: string;
  SubjectMatterExpert: string;
  Prework: string;
  RegistrationDate: string;
  StartTimeZoneCST: string;
  EndTimeZoneCST: string;
  StartTimeZoneEST: string;
  EndTimeZoneEST: string;
}

export const EventFields = [
  "Id",
  "Title",
  "EventDate",
  "EndDate",
  "Description",
  "Category",
  "CategoryColor",
  "TargetAudience",
  "Location",
  "Department",
  "Contact",
  "SubjectMatterExpert",
  "Prework",
  "RegistrationDate",
  "StartTimeZoneCST",
  "EndTimeZoneCST",
  "StartTimeZoneEST",
  "EndTimeZoneEST"
];
