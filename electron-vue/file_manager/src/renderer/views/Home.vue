<template>
    <div id="Home">
        <FIleManage/>
    </div>
</template>

<script>
    import FIleManage from "./FileManage";
    export default {
        name: "Home",
        components: {FIleManage}
    }
</script>

<style scoped>
    #Home {
        font-family: 'Avenir', Helvetica, Arial, sans-serif;
        -webkit-font-smoothing: antialiased;
        -moz-osx-font-smoothing: grayscale;
    }
</style>