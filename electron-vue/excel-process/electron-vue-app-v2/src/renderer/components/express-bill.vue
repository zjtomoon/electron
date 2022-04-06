<template>
  <div ref="imageDom" class="image-dom">
    <!-- <img src="../assets/kdd1473.jpg" alt=""> -->
    <div class="send-year">{{spitTime(textObj.sendTime).year}}</div>
    <div class="send-month">{{spitTime(textObj.sendTime).month}}</div>
    <div class="send-day">{{spitTime(textObj.sendTime).date}}</div>
    <div class="customer-code">{{textObj.customerCode}}</div>
    <div class="court">{{textObj.court}}</div>
    <div class="sender-phone">{{textObj.senderPhone}}</div>
    <div class="case-num">{{textObj.caseNum}}</div>
    <div class="court-time">{{textObj.courtTime}}</div>
    <div class="court-num">{{textObj.courtNum}}</div>
    <div class="res-notice">√</div>
    <div class="indictment">√</div>
    <div class="proof-notice">√</div>
    <div class="litigation-rights-inform">√</div>
    <div class="service-address-confirm">√</div>
    <div class="evidence-copy" v-if="textObj.evidenceCopy == '1'"><span>√</span>证据副本</div>
    <div class="small-claims-notice" v-if="textObj.smallClaimsNotice == '1'"><span>√</span>小额诉讼须知</div>
    <div class="summons">√</div>
    <div class="operator">{{textObj.operator}}</div>
    <div class="addressee">{{textObj.addressee}}</div>
    <div class="addressee-phone-1">{{textObj.addresseePhone1}}</div>
    <div class="addressee-phone-2" v-if="textObj.addresseePhone2">{{textObj.addresseePhone2}}</div>
    <div class="recipient-address">{{textObj.recipientAddress}}</div>
    <div class="r-s-office" v-if="textObj.rSoffice">{{textObj.rSoffice}}</div>
    <div class="sender-unit" v-if="textObj.senderUnit">{{textObj.senderUnit}}</div>
    <div class="sender-address" v-if="textObj.senderAddress">{{textObj.senderAddress}}</div>
  </div>
</template>

<script>
import html2canvas from 'html2canvas'
const fs = require('fs')
const xlsx = require('node-xlsx')

export default {
  props: {
    log: Object
  },
  data () {
    return {
      textObj: {
        sendTime: 42830,
        customerCode: 'PY29',
        court: '立案庭',
        senderPhone: '020-69122573',
        caseNum: '2020粤0113民初5564',
        courtTime: '20.6.5下15:00',
        courtNum: '十三庭',
        evidenceCopy: '1',
        smallClaimsNotice: '1',
        operator: '张伟',
        addressee: '李易真',
        addresseePhone1: '13728030495',
        addresseePhone2: '13728030495',
        recipientAddress: '广东省广州市海珠区蟠龙东路98号202房',
        rSoffice: '番禺桥南',
        senderUnit: '广州市番禺区',
        senderAddress: '广州市番禺区市桥桥兴大道733号'
      },
      sourcePath: '',
      folderPath: '',
      fileNum: 0
    }
  },
  methods: {
    async start () {
      const keyOrderMap = {}
      Object.keys(this.textObj).forEach((x, i) => {
        keyOrderMap[i] = x
      })
      this.log.text = `读取文件...`
      setTimeout(async () => {
        const buffer = fs.readFileSync(this.sourcePath)
        this.log.text = `正在解析...`
        const workSheetsFromList = xlsx.parse(buffer)
        const sheetData = workSheetsFromList[0].data
        // 去掉空行
        for (let i = sheetData.length - 1; i >= 0; i--) {
          if (!sheetData[i].length || sheetData[i].every(x => !x)) {
            sheetData.splice(i, 1)
          }
        }
        for (let i = 1; i < sheetData.length; i++) {
          for (let j = 0; j < sheetData[0].length; j++) {
            this.textObj[keyOrderMap[j]] = sheetData[i][j]
          }
          await this.sleep(1)
          await this.generatePicture()
        }
        this.log.text = `处理结束！共生成${this.fileNum}张图片`
        this.fileNum = 0
        this.$emit('end')
      }, 1)
    },
    sleep (n) {
      return new Promise((resolve, reject) => {
        setTimeout(() => {
          resolve()
        }, n)
      })
    },
    spitTime (excelTimeNum) {
      if (!excelTimeNum) {
        return {}
      }
      const time = new Date(((excelTimeNum - 19 - 70 * 365) * 86400 - 8 * 3600) * 1000)
      return {
        year: time.getFullYear(),
        month: time.getMonth() + 1,
        date: time.getDate()
      }
    },
    /**
     * 将页面指定节点内容转为图片
     * 1.拿到想要转换为图片的内容节点DOM；
     * 2.转换，拿到转换后的canvas
     * 3.转换为图片
     */
    generatePicture () {
      const imageDom = this.$refs.imageDom
      const opt = {
        width: imageDom.offsetWidth,
        height: imageDom.offsetHeight,
        scale: 1,
        backgroundColor: null
      }
      return html2canvas(imageDom, opt).then(canvas => {
        // 转成图片，生成图片地址
        const base64Data = canvas.toDataURL().replace(/^data:image\/\w+;base64,/, '')
        const dataBuffer = Buffer.from(base64Data, 'base64')
        // 生成文件夹
        const folderName = '快递单模板图'
        const caseNum = this.textObj.caseNum.replace(/2.+?初/, '')
        const addressee = this.textObj.addressee
        let filename = `${caseNum}${addressee}.png`
        let path = `${this.folderPath}/${folderName}/${filename}`
        if (!fs.existsSync(`${this.folderPath}/${folderName}`)) {
          fs.mkdirSync(`${this.folderPath}/${folderName}`)
        }
        let index = 1
        while (fs.existsSync(path)) {
          filename = `${caseNum}${addressee + index++}.png`
          path = `${this.folderPath}/${folderName}/${filename}`
        }
        fs.writeFileSync(path, dataBuffer)
        this.log.text = `成功生成${++this.fileNum}张图片`
      })
    }
  }
}
</script>

<style lang="scss" scoped>
.image-dom {
  position: fixed;
  top: -5000px;
  width: 1015px;
  height: 560px;
  font-weight: bold;
  font-size: 14px;
  cursor: pointer;
  .send-year {
    position: absolute;
    left: 306px;
    top: 109px;
  }
  .send-month {
    position: absolute;
    left: 355px;
    top: 109px;
  }
  .send-day {
    position: absolute;
    left: 391px;
    top: 109px;
  }
  .customer-code {
    position: absolute;
    left: 414px;
    top: 147px;
  }
  .court {
    position: absolute;
    left: 289px;
    top: 164px;
  }
  .sender-phone {
    position: absolute;
    left: 391px;
    top: 200px;
  }
  .case-num {
    position: absolute;
    left: 162px;
    top: 224px;
  }
  .court-time {
    position: absolute;
    font-size: 12px;
    left: 400px;
    top: 225px;
  }
  .court-num {
    position: absolute;
    font-size: 12px;
    left: 427px;
    top: 240px;
  }
  .res-notice{
    position: absolute;
    font-size: 12px;
    font-weight: bold;
    left: 95px;
    top: 252px;
  }
  .indictment{
    position: absolute;
    font-size: 12px;
    font-weight: bold;
    left: 95px;
    top: 263px;
  }
  .proof-notice{
    position: absolute;
    font-size: 12px;
    font-weight: bold;
    left: 95px;
    top: 293px;
  }
  .litigation-rights-inform{
    position: absolute;
    font-size: 12px;
    font-weight: bold;
    left: 95px;
    top: 303px;
  }
  .service-address-confirm{
    position: absolute;
    font-size: 12px;
    font-weight: bold;
    left: 95px;
    top: 323px;
  }
  .evidence-copy{
    position: absolute;
    font-size: 12px;
    left: 214px;
    top: 302px;
    span {
      margin-right: 3px;
      transform: translateY(1px);
      display: inline-block;
    }
  }
  .small-claims-notice{
    position: absolute;
    font-size: 12px;
    left: 214px;
    top: 322px;
    span {
      margin-right: 3px;
      transform: translateY(1px);
      display: inline-block;
    }
  }
  .summons{
    position: absolute;
    font-size: 12px;
    left: 360px;
    top: 242px;
  }
  .operator{
    position: absolute;
    left: 395px;
    bottom: 187px;
  }
  .addressee{
    position: absolute;
    left: 528px;
    top: 165px;
  }
  .addressee-phone-1{
    position: absolute;
    left: 753px;
    top: 142px;
  }
  .addressee-phone-2{
    position: absolute;
    left: 753px;
    top: 161px;
  }
  .recipient-address{
    position: absolute;
    left: 549px;
    top: 240px;
    width: 350px;
    word-break: break-all;
  }
  .r-s-office {
    position: absolute;
    left: 137px;
    top: 106px;
  }
  .sender-unit {
    position: absolute;
    left: 105px;
    top: 164px;
  }
  .sender-address {
    position: absolute;
    left: 126px;
    top: 197px;
  }
}
</style>