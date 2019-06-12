export class MockupData{
  public static getJuneData():any[]{
    return [
      {
        title: 'All Day June Event',
        start: '2019-06-11',
        end: '2019-06-11'
      },
      {
        title: 'Long June Event',
        start: '2019-05-07',
        end: '2019-05-10'
      }
    ];
  }
  public static getMayData(): any[] {
    return [
      {
        title: 'All Day Event',
        start: '2019-05-01',
        end: '2019-05-01'
      },
      {
        title: 'Long Event',
        start: '2019-05-07',
        end: '2019-05-10'
      },
      {
        id: 999,
        title: 'Repeating Event',
        start: '2019-05-09T16:00:00'
      },
      {
        id: 999,
        title: 'Repeating Event',
        start: '2019-05-16T16:00:00'
      },
      {
        title: 'Conference',
        start: '2019-05-11',
        end: '2019-05-13'
      },
      {
        title: 'Meeting',
        start: '2019-05-12T10:30:00',
        end: '2019-05-12T12:30:00'
      },
      {
        title: 'Lunch',
        start: '2019-05-12T12:00:00'
      },
      {
        title: 'Meeting',
        start: '2019-05-12T14:30:00'
      },
      {
        title: 'Happy Hour',
        start: '2019-05-12T17:30:00'
      },
      {
        title: 'Dinner',
        start: '2019-05-12T20:00:00'
      },
      {
        title: 'Birthday Party',
        start: '2019-05-13T07:00:00'
      },
      {
        title: 'Click for Google',
        url: 'http://google.com/',
        start: '2019-05-28'
      }
    ];
  }
}
