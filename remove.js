function clean_all_events() {
    const thisYear = new Date().getFullYear();
    const calendar = CalendarApp.getCalendarById('gm.nuu.edu.tw_853lo4f1bcijuhk8h1kkcp6128@group.calendar.google.com');
    const events = calendar.getEvents(
        new Date([thisYear, '8/1'].join('/')), 
        new Date([thisYear + 1, '8/1'].join('/')));
    events.forEach(event => event.deleteEvent());
}