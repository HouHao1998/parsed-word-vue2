<template>
  <v-app>
    <v-main>
      <v-container>
        <v-row>
          <v-col cols="12" class="text-center">
            <h1>中国标准化杂志社-格式教研系统</h1>
          </v-col>
        </v-row>
        <file-upload></file-upload>
      </v-container>
    </v-main>
  </v-app>
</template>

<script>
import FileUpload from './components/FileUpload.vue';

export default {
  components: {
    FileUpload
  }
};
</script>

<style>
h1 {
  margin-top: 20px;
}
</style>
