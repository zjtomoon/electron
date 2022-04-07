import Vue from 'vue'
import Router from 'vue-router'

Vue.use(Router)

export default new Router({
  routes: [
    {
      path: '/Home',
      name: 'Home',
      component: Home
    },
    {
      path: '/',
      name: 'Login',
      // eslint-disable-next-line no-undef
      component: Login
    },
    {
      path: '/Register',
      name: 'Register',
      // eslint-disable-next-line no-undef
      component: Register
    }

  ]
})
