<template>
  <v-container>
    <v-row justify="center">
      <v-col cols="12" sm="8">
        <h2>解析内容</h2>
        <div v-if="content">
          <p><strong>中文题名:</strong> {{ content["中文题名"] }}</p>
          <p><strong>作者姓名:</strong> {{ content["作者姓名"] }}</p>
          <p><strong>工作单位及通信方式:</strong> {{ content["工作单位及通信方式"] }}</p>
          <p><strong>中文摘要引题:</strong> {{ content["中文摘要"]["引题"] }}</p>
          <p><strong>中文摘要内容:</strong> {{ content["中文摘要"]["内容"] }}</p>
          <p><strong>中文关键词引题:</strong> {{ content["中文关键词"]["引题"] }}</p>
          <p><strong>中文关键词内容:</strong> {{ content["中文关键词"]["内容"] }}</p>
          <p><strong>英文题名:</strong> {{ content["英文题名"] }}</p>
          <p><strong>英文作者姓名:</strong> {{ content["英文作者姓名"] }}</p>
          <p><strong>英文工作单位及通信方式:</strong> {{ content["英文工作单位及通信方式"] }}</p>
          <p><strong>英文摘要引题:</strong> {{ content["英文摘要"]["引题"] }}</p>
          <p><strong>英文摘要内容:</strong> {{ content["英文摘要"]["内容"] }}</p>
          <p><strong>英文关键词引题:</strong> {{ content["英文关键词"]["引题"] }}</p>
          <p><strong>英文关键词内容:</strong> {{ content["英文关键词"]["内容"] }}</p>
          <h3>章节编号和标题</h3>
          <ul>
            <li v-for="chapter in content['章节编号和标题']" :key="chapter">{{ chapter }}</li>
          </ul>
          <h3>节编号和标题</h3>
          <ul>
            <li v-for="section in content['节编号和标题']" :key="section">{{ section }}</li>
          </ul>
          <h3>正文内容</h3>
          <ul>
            <li v-for="text in content['正文内容']" :key="text">{{ text }}</li>
          </ul>
          <h3>表格</h3>
          <ul>
            <li v-for="table in content['表格']" :key="table['编号和标题']">
              <strong>编号和标题:</strong> {{ table["编号和标题"] }}
              <strong>内容:</strong> {{ table["内容"] }}
              <strong>表注:</strong> {{ table["表注"] }}
            </li>
          </ul>
          <h3>参考文献</h3>
          <ul>
            <li v-for="ref in content['参考文献']" :key="ref">{{ ref }}</li>
          </ul>
        </div>
        <div v-else>
          <p>解析内容将显示在这里。</p>
        </div>
      </v-col>
    </v-row>
  </v-container>
</template>

<script>
export default {
  props: {
    content: Object,
  },
};
</script>
