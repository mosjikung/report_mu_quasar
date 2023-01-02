const routes = [
  {
    path: "/main",
    component: () => import("layouts/MainLayout.vue"),
    children: [{ path: "", component: () => import("pages/Index.vue") }],
  },
  {
    path: "/test",
    component: () => import("pages/Test.vue"),
    children: [{ path: "", component: () => import("pages/Test.vue") }],
  },
  {
    path: "/test2",
    component: () => import("pages/Test2.vue"),
    children: [{ path: "", component: () => import("pages/Test2.vue") }],
  },
  {
    path: "/",
    component: () => import("pages/Login.vue"),
    children: [
      { path: "", component: () => import("pages/Login.vue") },
      { path: "/login", component: () => import("pages/Login.vue") },
    ],
  },

  // Always leave this as last one,
  // but you can also remove it
  {
    path: "/:catchAll(.*)*",
    component: () => import("pages/Error404.vue"),
  },
];

export default routes;
