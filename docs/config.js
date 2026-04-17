window.APP_CONFIG = {
  TIMEZONE: "Asia/Tokyo",
  START_HOUR: 9,
  END_HOUR: 21,
  CORS_PROXY: "https://api.allorigins.win/raw?url={url}",
  CORS_PROXIES: [
    "https://api.allorigins.win/raw?url={url}",
    "https://cors.isomorphic-git.org/{url}",
    "https://corsproxy.io/?url={url}"
  ],
  TEMPLATE_PATH: "./template.xlsx",
  VENUES: [
    {
      id: "gym-group-calendar",
      name: "东体育馆",
      embedUrls: [
        "https://calendar.google.com/calendar/u/0/embed?src=cjf3gmhfu52jvshevng5ibslok@group.calendar.google.com&ctz=Asia/Tokyo"
      ]
    },
    {
      id: "gym-sports-center",
      name: "西体育馆",
      embedUrls: [
        "https://calendar.google.com/calendar/u/0/embed?src=hu.sportscenter@gmail.com&ctz=Asia/Tokyo&pli=1"
      ]
    }
  ]
};
