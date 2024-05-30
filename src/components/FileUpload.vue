<template>
  <v-container>
    <v-row justify="center" class="mt-5" v-if="!file">
      <v-col cols="12" sm="8" class="text-center">
        <div class="upload-box" @click="triggerFileInput">
          <v-icon size="48" color="primary">mdi-upload</v-icon>
          <p>请上传你需要解析的word文档</p>
          <input
              type="file"
              ref="fileInput"
              @change="onFileChange"
              accept=".docx"
              style="display: none;"
          />
        </div>
      </v-col>
    </v-row>
    <v-row justify="center" v-if="!file">
      <v-col cols="12" sm="8" class="text-center">
        <v-btn color="info" href="http://jiangyakeji.com" target="_blank">
          AI智能聊天助手chatGPT
        </v-btn>
      </v-col>
    </v-row>
    <v-row v-else class="mt-3 align-center">
      <v-col cols="auto" class="d-flex align-center">
        <v-icon size="36" color="blue darken-2" class="mr-2">mdi-file-word-box</v-icon>
        <h2 class="mb-0 mr-2">{{ file.name }}</h2>
        <v-btn color="secondary" @click="replaceFile" class="mr-2">替换文件</v-btn>
        <v-btn color="primary" @click="parseFile" class="mr-2">解析</v-btn>
        <v-btn color="success" @click="formatFile">格式化</v-btn>
        <input
            type="file"
            ref="replaceFileInput"
            @change="onFileChange"
            accept=".docx"
            style="display: none;"
        />
      </v-col>
      <v-col cols="auto" class="text-right ml-auto">
        <v-btn color="info" href="http://jiangyakeji.com" target="_blank">
          AI智能聊天助手chatGPT
        </v-btn>
      </v-col>
    </v-row>
    <v-row v-if="file" class="mt-5">
      <v-col cols="12" sm="6" class="preview-container">
        <h3>文档预览</h3>
        <vue-office-docx :src="src" class="preview"></vue-office-docx>
      </v-col>
      <v-col cols="12" sm="6">
        <h3>解析结果</h3>
        <div class="results">
          <div v-for="item in parsedContent" :key="item.内容" class="section">
            <p><strong>{{ item.类型 }}:</strong></p>
            <p :style="getStyle(item)">{{ checkFormat(item) }}</p>
            <p>{{ item.内容 }}</p>
            <hr>
          </div>
        </div>
      </v-col>
    </v-row>
  </v-container>
</template>

<script>
import VueOfficeDocx from '@vue-office/docx';
import '@vue-office/docx/lib/index.css';
import axios from 'axios';

export default {
  components: {
    VueOfficeDocx
  },
  data() {
    return {
      file: null,
      src: '',
      parsedContent: []
    };
  },
  methods: {
    triggerFileInput() {
      this.$refs.fileInput.click();
    },
    onFileChange(event) {
      if (!event.target.files.length) {
        console.log("No file selected!");
        return;
      }
      this.file = event.target.files[0];
      this.readFile(this.file);
    },
    readFile(file) {
      const fileReader = new FileReader();
      fileReader.readAsArrayBuffer(file);
      fileReader.onload = () => {
        this.src = fileReader.result;
      };
      fileReader.onerror = (error) => {
        console.error("Error reading file:", error);
      };
    },
    parseFile() {
      if (!this.file) {
        alert('请选择一个文件');
        return;
      }
      const formData = new FormData();
      formData.append('file', this.file);

      axios.post('http://127.0.0.1:5000/upload', formData, {
        headers: {
          'Content-Type': 'multipart/form-data'
        }
      }).then(response => {
        this.parsedContent = response.data;
      }).catch(error => {
        console.error(error);
      });
    },
    formatFile() {
      if (!this.file) {
        alert('请选择一个文件');
        return;
      }
      const formData = new FormData();
      formData.append('file', this.file);

      axios.post('http://127.0.0.1:5000/format', formData, {
        headers: {
          'Content-Type': 'multipart/form-data'
        },
        responseType: 'blob' // 确保服务器返回的是Blob对象
      }).then(response => {
        // 创建一个链接指向返回的Blob对象，并自动触发下载
        const url = window.URL.createObjectURL(new Blob([response.data]));
        const link = document.createElement('a');
        link.href = url;
        link.setAttribute('download', '格式化的文档.docx');
        document.body.appendChild(link);
        link.click();
      }).catch(error => {
        console.error(error);
      });
    },
    replaceFile() {
      this.$refs.replaceFileInput.click();
    },
    checkFormat(item) {
      const formatRequirements = {
        "中文题名": { font: "黑体", size: 19 },
        "作者姓名": { font: "楷体", size: 12.5 },
        "工作单位及通信方式": { font: "宋体", size: 9 },
        "中文摘要引题": { font: "黑体", size: 9 },
        "中文摘要内容": { font: "仿宋", size: 9 },
        "中文关键词引题": { font: "黑体", size: 9 },
        "中文关键词内容": { font: "仿宋", size: 9 },
        "英文题名": { font: "黑体", size: 14 },
        "英文作者姓名": { font: "宋体", size: 11 },
        "英文工作单位及通信方式": { font: "宋体", size: 9 },
        "英文摘要引题": { font: "黑体", size: 9 },
        "英文摘要内容": { font: "宋体", size: 9 },
        "英文关键词引题": { font: "黑体", size: 9 },
        "英文关键词内容": { font: "宋体", size: 9 },
        "章节编号和标题": { font: "黑体", size: 14 },
        "节编号和标题": { font: "黑体", size: 11 },
        "正文内容": { font: "宋体", size: 11 },
        "插图、表格编号和标题": { font: "黑体", size: 9 },
        "表格内容、表注和图注": { font: "宋体", size: 9 },
        "致谢引题": { font: "黑体", size: 11 },
        "致谢内容": { font: "楷体", size: 11 },
        "参考文献引题": { font: "黑体", size: 14 },
        "参考文献内容": { font: "宋体", size: 9 },
        "附录编号和标题": { font: "黑体", size: 14 },
        "附录内容": { font: "宋体", size: 9 }
      };

      const requirement = formatRequirements[item.类型];
      if (!requirement) {
        return "无特定要求";
      }

      if (item.字体 instanceof Array && item.字号 instanceof Array) {
        const fontMatch = item.字体.every(f => f === requirement.font);
        const sizeMatch = item.字号.every(s => s === requirement.size);
        return fontMatch && sizeMatch ? "文字样式正确" : "文字样式错误";
      }

      const fontMatch = item.字体 === requirement.font;
      const sizeMatch = item.字号 === requirement.size;

      return fontMatch && sizeMatch ? "文字样式正确" : "文字样式错误";
    },
    getStyle(item) {
      return {
        color: this.checkFormat(item) === "文字样式正确" ? 'green' : 'red'
      };
    }
  }
};
</script>

<style>
.upload-box {
  border: 2px dashed #ccc;
  border-radius: 10px;
  padding: 40px;
  text-align: center;
  position: relative;
  cursor: pointer;
  margin: 0 auto;
}
.upload-box v-icon {
  font-size: 48px;
  color: #1976d2;
}
.upload-box p {
  margin: 20px 0 0;
  color: #777;
  font-size: 16px;
}
.preview-container {
  padding-left: 10px; /* 减小预览框与左侧页边距的距离 */
}
.preview, .results {
  border: 1px solid #ddd;
  border-radius: 10px;
  padding: 15px;
  height: auto;
  overflow-y: auto;
  width: 100%;
  max-height: 80vh; /* 设置最大高度为视窗高度的80% */
  min-height: 500px;
}
.section {
  margin-bottom: 10px;
}
hr {
  border-top: 1px dashed #ddd;
}
.text-right {
  text-align: right;
}
</style>
