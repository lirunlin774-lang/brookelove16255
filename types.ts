
export interface ScheduleItem {
  id: string;
  time: string;
  content: string;
  speaker: string;
}

export interface ExpenseItem {
  id: string;
  category: string;
  project: string;
  price: number;
  unit: string;
  quantity: number;
  total: number;
  description: string;
}

export interface ActivityData {
  channelName: string;
  activityDate: string; // YYYY-MM-DD
  startTime: string;
  endTime: string;
  location: string;
  participantsDesc: string;
  submitDate: string;
  schedule: ScheduleItem[];
  participantCount: number;
  expenses: ExpenseItem[];
}
