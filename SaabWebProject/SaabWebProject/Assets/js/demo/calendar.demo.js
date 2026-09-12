/*
Template Name: HUD - Responsive Bootstrap 5 Admin Template
Version: 1.6.0
Author: Sean Ngu
Website: http://www.seantheme.com/hud/
*/

var handleRenderFullcalendar = function () {
  // external events
  var containerEl = document.getElementById("external-events");
  var Draggable = FullCalendarInteraction.Draggable;
  new Draggable(containerEl, {
    itemSelector: ".fc-event-link",
    eventData: function (eventEl) {
      return {
        title: eventEl.innerText,
        color: eventEl.getAttribute("data-color"),
      };
    },
  });

  // fullcalendar
  var d = new Date();
  var month = d.getMonth() + 1;
  month = month < 10 ? "0" + month : month;
  var year = d.getFullYear();
  var day = d.getDate();
  var today = moment().startOf("day");
  var calendarElm = document.getElementById("calendar");
  var calendar = new FullCalendar.Calendar(calendarElm, {
    headerToolbar: {
      left: "dayGridMonth,timeGridWeek,timeGridDay",
      center: "title",
      right: "prev,next today",
    },
    locale: "fa",
	isRTL: true,
    buttonText: {
      today: "امروز",
      month: "ماه",
      week: "هفته",
      day: "روز",
    },
    initialView: "dayGridMonth",
    editable: true,
    droppable: true,
    themeSystem: "bootstrap",
    eventLimit: true, // for all non-TimeGrid views
    views: {
      timeGrid: {
        eventLimit: 6, // adjust to 6 only for timeGridWeek/timeGridDay
      },
    },
    events: [
      {
        title: "سفر به لندن",
        start: year + "-" + month + "-01",
        end: year + "-" + month + "-05",
        color: app.color.theme,
      },
      {
        title: "ملاقات با شان نگو",
        start: year + "-" + month + "-02T06:00:00",
        color: app.color.blue,
      },
      {
        title: "برنامه های موبایل طوفان مغزی",
        start: year + "-" + month + "-10",
        end: year + "-" + month + "-12",
        color: app.color.pink,
      },
      {
        title: "  دیدن قلعه در آکسفورد",
        start: year + "-" + month + "-05T08:45:00",
        end: year + "-" + month + "-06T18:00",
        color: app.color.indigo,
      },
      {
        title: "سفر پاریس",
        start: year + "-" + month + "-12",
        end: year + "-" + month + "-16",
      },
      {
        title: "نام دامنه  ",
        start: year + "-" + month + "-15",
        end: year + "-" + month + "-15",
        color: app.color.blue,
      },
      {
        title: "سفر کمبریج",
        start: year + "-" + month + "-19",
        end: year + "-" + month + "-19",
      },
      {
        title: "شرکت اپل را ببینید",
        start: year + "-" + month + "-22T05:00:00",
        color: app.color.green,
      },
      {
        title: "کلاس تمرین",
        start: year + "-" + month + "-22T07:30:00",
        color: app.color.orange,
      },
      {
        title: "ضبط زنده",
        start: year + "-" + month + "-22T03:00:00",
        color: app.color.blue,
      },
      {
        title: "اعلان",
        start: year + "-" + month + "-22T15:00:00",
        color: app.color.red,
      },
      {
        title: "شام",
        start: year + "-" + month + "-22T18:00:00",
      },
      {
        title: "بحث جدید آندروید برنامه",
        start: year + "-" + month + "-25T08:00:00",
        end: year + "-" + month + "-25T10:00:00",
        color: app.color.red,
      },
      {
        title: "ارائه طرح بازاریابی",
        start: year + "-" + month + "-25T12:00:00",
        end: year + "-" + month + "-25T14:00:00",
        color: app.color.blue,
      },
      {
        title: "تعقیب کردن",
        start: year + "-" + month + "-26T12:00:00",
        color: app.color.orange,
      },
      {
        title: "گرگ",
        start: year + "-" + month + "-26T08:00:00",
        color: app.color.orange,
      },
      {
        title: "ناهار با ریچارد",
        start: year + "-" + month + "-28T14:00:00",
        color: app.color.blue,
      },
      {
        title: "میزبانی وب  ",
        start: year + "-" + month + "-30",
        color: app.color.blue,
      },
    ],
  });

  calendar.render();
};

/* Controller
------------------------------------------------ */
$(document).ready(function () {
  handleRenderFullcalendar();
});
