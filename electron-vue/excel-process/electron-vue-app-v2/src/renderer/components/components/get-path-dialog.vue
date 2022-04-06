<template>
  <el-dialog
    custom-class="get-path-dialog"
    title="选择下载位置"
    :visible="showDialog"
    width="470px"
    @close="close"
    @open="open">
    <el-upload
      class="el-upload-folder"
      ref="elUploadFolder"
      drag
      action=""
      :limit="1"
      :auto-upload="false"
      :on-change="onFolderChange"
      :on-remove="onFolderRemove"
      list-type="picture">
      <i class="el-icon-upload"></i>
      <div class="el-upload__text">将文件夹拖到此处</div>
    </el-upload>
    <span slot="footer" class="dialog-footer">
      <el-button @click="showDialog = false">取 消</el-button>
      <el-button type="primary" @click="toDetermine">确 定</el-button>
    </span>
  </el-dialog>
</template>

<script>
import { Dialog, Button, Upload, Message } from 'element-ui'
const { shell } = require('electron')
const os = require('os')

export default {
  components: {
    [Button.name]: Button,
    [Dialog.name]: Dialog,
    [Upload.name]: Upload
  },
  data () {
    return {
      showDialog: false,
      folderPath: ''
    }
  },
  methods: {
    showItemInFolder () {
      shell.showItemInFolder(os.homedir())
    },
    onFolderChange (file, fileList) {
      fileList[0].url = require('@/assets/icon-folder.png')
      this.folderPath = file.raw.path
    },
    onFolderRemove (file, fileList) {
      this.folderPath = ''
    },
    close () {
      this.$refs.elUploadFolder.clearFiles()
      this.folderPath = ''
      this.showDialog = false
    },
    open () {
      this.$nextTick(() => {
        this.$refs.elUploadFolder.$refs['upload-inner'].$refs.input.onclick = (e) => {
          e.preventDefault()
          this.showItemInFolder()
        }
      })
    },
    toDetermine () {
      if (!this.folderPath) {
        Message.warning({
          message: '请选择文件夹',
          center: true
        })
        return
      }
      this.$emit('toDetermine', this.folderPath)
    }
  }
}
</script>

<style lang="scss">
.get-path-dialog {
  .el-dialog__body {
    padding-bottom: 0;
  }
  .el-dialog__footer {
    text-align: center;
  }
}
</style>

<style lang="scss" scoped>
.el-upload-folder {
  display: flex;
  flex-direction: column;
  align-items: center;
  min-height: 220px;
  /deep/.el-upload-list__item {
    height: auto;
    padding: 0 10px 0 90px;
    display: flex;
    align-items: center;
    .el-upload-list__item-thumbnail {
      height: 15px;
      width: 15px;
    }
    .el-icon-document {
      display: none;
    }
    .el-upload-list__item-name {
      margin: 0;
      margin-left: 8px;
      padding: 0;
    }
    .el-icon-close {
      top: 50%;
      transform: translateY(-50%);
    }
  }
}
</style>