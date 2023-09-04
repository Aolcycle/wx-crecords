<template>
  <div>
    <Spin :spinning="spinning" tip="此过程大约需要2-5分钟,请耐心等候...">
      <div style="width: 200px; height: 200px">
        <Upload
          :file-list="fileList"
          :max-count="1"
          :custom-request="handleImport"
          accept=".xls, .xlsx, .csv"
          @remove="handleRemove"
        >
          <AButton> 选择文件 </AButton>
        </Upload>
        <div>
          <AButton @click="onLength">分析数据</AButton>
          <AButton @click="handleExport">导出</AButton>
        </div>
      </div>
      <div>
        <ATable :data-source="dataSource" :columns="columns" :scroll="{ y: 550 }">
          <template #emptyText> 暂无数据 </template>
        </ATable>
      </div>
    </Spin>
  </div>
</template>

<script lang="ts">
import { defineComponent } from 'vue'
import { Upload, Button, Spin, Table } from 'ant-design-vue'
import dayjs from 'dayjs'

export default defineComponent({
  name: 'App',
  components: {
    Upload,
    AButton: Button,
    Spin,
    ATable: Table
  }
})
</script>

<script setup lang="ts">
import { ref } from 'vue'
import type { UploadProps } from 'ant-design-vue'
import { read, utils } from 'xlsx'
import { exportExcel } from './utils/excelUtil'

const dataList = [] as DataType[]

const dataList_type_1 = []

const onLength = () => {
  let closestTime = 0
  let closestDate = ''
  let closesContent = ''
  let smallestDifference = 24 * 60 * 60 * 1000 // 初始化为一天的毫秒数
  const targetTime = '04:00:00'
  let isSend_1 = 0
  let isSend_0 = 0
  let type_10000 = 0
  let wenzi = 0
  let yuyin = 0
  let biaoqing = 0

  let earliestTimestamp = Number.MAX_SAFE_INTEGER
  let earliestDate = ''
  let earliestContent = ''

  dataSource.value.forEach((item) => {
    if (item.talker === 'wxid_cjsjytzw4wms22') {
      dataList.push({
        msgId: item.msgId,
        talker: item.talker,
        createTime: item.createTime,
        type: item.type,
        isSend: item.isSend,
        content: item.content
      })

      if (item.__createTime__ < earliestTimestamp) {
        earliestTimestamp = item.__createTime__
        earliestContent = item.content
        earliestDate = item.createTime
      }

      if (item.isSend == '0') {
        isSend_0++
      } else {
        isSend_1++
      }
      if (item.type == '10000') {
        type_10000++
      }
      if (item.type == '1') {
        try {
          wenzi = wenzi + (item.content.length ? item.content.length : 0)
        } catch (e) {}

        dataList_type_1.push({
          msgId: item.msgId,
          talker: item.talker,
          createTime: item.createTime,
          type: item.type,
          isSend: item.isSend,
          content: item.content
        })
      }
      if (item.type == '34') {
        yuyin++
      }
      if (item.type == '47') {
        biaoqing++
      }

      const time = item.createTime.split(' ')[1] // 提取时间部分
      const timeObj = dayjs()
        .hour(parseInt(time.split(':')[0]))
        .minute(parseInt(time.split(':')[1]))
        .second(parseInt(time.split(':')[2]))
      const targetTimeObj = dayjs()
        .hour(parseInt(targetTime.split(':')[0]))
        .minute(parseInt(targetTime.split(':')[1]))
        .second(parseInt(targetTime.split(':')[2]))

      let difference = Math.abs(timeObj.diff(targetTimeObj, 'millisecond'))

      // 如果时间跨越午夜（00:00:00），需要考虑这一点
      if (difference > 12 * 60 * 60 * 1000) {
        difference = 24 * 60 * 60 * 1000 - difference
      }

      if (difference < smallestDifference) {
        smallestDifference = difference
        closestTime = time
        closestDate = item.createTime
        closesContent = item.content
      }
    }
  })
  console.log(
    '我们一共聊了：',
    dataList.length,
    '条, 其中我发的消息有',
    isSend_1,
    '条, 你发的消息有',
    isSend_0,
    '条, 我发消息的占比是',
    (isSend_1 / (isSend_1 + isSend_0)) * 100
  )
  console.log('撤回消息了：', type_10000, '条')
  console.log('联系人最早一条信息的时间是:', earliestTimestamp, earliestDate)
  console.log('联系人最早一条信息的内容是:', earliestContent)
  const date1 = dayjs(dataList[0].createTime).format('YYYY-MM-DD')
  console.log(
    '最晚的聊天的时间是：',
    closestTime,
    '那天是：' + closestDate,
    '内容是：',
    closesContent
  )
  console.log('只看文字消息，我们一共聊了：' + wenzi + '字')
  console.log('我们一共发了：' + yuyin + '条语音消息，表情包发了' + biaoqing + '条')
}

// table列表数据
const dataSource = ref<DataType[]>()
type FiltersType = {
  text: string
  value: string
}
const talkerList = ref<FiltersType[]>([])
const typeList = ref<FiltersType[]>([
  {
    text: '文本内容',
    value: '1'
  },
  {
    text: '位置信息',
    value: '2'
  },
  {
    text: '图片及视频',
    value: '3'
  },
  {
    text: '语音消息',
    value: '34'
  },
  {
    text: '名片（公众号名片）',
    value: '42'
  },
  {
    text: '表情包',
    value: '47'
  },
  {
    text: '定位信息',
    value: '48'
  },
  {
    text: '小程序链接',
    value: '49'
  },
  {
    text: '撤回消息提醒',
    value: '10000'
  },
  {
    text: '照片',
    value: '1048625'
  },
  {
    text: '链接',
    value: '16777265'
  },
  {
    text: '文件',
    value: '285212721'
  },
  {
    text: '微信转账',
    value: '419430449'
  },
  {
    text: '微信红包1',
    value: '436207665'
  },
  {
    text: '微信红包2',
    value: '469762097'
  }
])

interface DataType {
  msgId: string
  talker: string
  createTime: string
  content: string
  type: string
  isSend: string
  __createTime__?: string
}
interface ColumnsType {
  title: string
  key: string
  align: string
  ellipsis?: boolean
  width?: number
  dataIndex?: string
  filters?: FiltersType[]
  onFilter?: (value: string, record: Record<string, any>) => boolean
  filterSearch?: (input: string, record: Record<string, any>) => boolean
}
// 表格header
const columns = ref<ColumnsType[]>([
  {
    title: '聊天顺序编号',
    key: 'msgId',
    dataIndex: 'msgId',
    width: 90,
    align: 'center'
  },
  {
    title: '聊天用户Id',
    dataIndex: 'talker',
    key: 'talker',
    width: 120,
    align: 'center',
    filters: talkerList.value,
    onFilter: (value, record) => record.talker.indexOf(value),
    filterSearch: (input, filter) => (filter.value as string).indexOf(input) > -1
  },
  {
    title: '聊天时间',
    dataIndex: 'createTime',
    key: 'createTime',
    width: 90,
    align: 'center'
  },
  {
    title: '聊天内容类型',
    dataIndex: 'type',
    key: 'type',
    width: 90,
    align: 'center',
    filters: typeList.value,
    onFilter: (value, record) => record.type.startsWith(value)
  },
  {
    title: '标识消息发送者(1自己0对方)',
    dataIndex: 'isSend',
    key: 'isSend',
    width: 90,
    align: 'center',
    filters: [
      {
        text: '自己',
        value: '1'
      },
      {
        text: '对方',
        value: '0'
      }
    ],
    onFilter: (value, record) => record.type.startsWith(value)
  },
  {
    title: '聊天内容',
    dataIndex: 'content',
    key: 'content',
    ellipsis: true,
    width: 150,
    align: 'left'
  }
])

const spinning = ref<boolean>(false)

const fileList = ref<UploadProps['fileList']>([])

const handleImport = (file): void => {
  spinning.value = true
  fileList.value = [...(fileList.value || []), file.file]
  const reader: FileReader = new FileReader()
  reader.readAsBinaryString(file.file)
  reader.onload = (ev): void => {
    if (ev.target) {
      const res = ev.target.result
      const worker = read(res, { type: 'binary' })
      // 将返回的数据转换为json对象的数据
      const jsonData = utils.sheet_to_json(worker.Sheets[worker.SheetNames[0]])
      spinning.value = false

      const newTalker = new Set<string>()
      jsonData.forEach((item: DataType) => {
        newTalker.add(item.talker)
        item.__createTime__ = item.createTime
        item.createTime = dayjs(item.createTime).format('YYYY-MM-DD HH:mm:ss')
      })

      dataSource.value = jsonData

      newTalker.forEach((item) => {
        talkerList.value.push({
          text: item,
          value: item
        })
      })

      columns.value[1].filters = talkerList.value
    }
  }
}

const handleExport = (): void => {
  const lists = ref<DataType[]>(dataList_type_1)
  const titleArr = ['msgId', 'talker', 'createTime', 'type', 'isSend', 'content']
  exportExcel(lists.value as [], '聊天记录', titleArr, 'sheet1')
}

const handleRemove: UploadProps['onRemove'] = (file) => {
  if (fileList.value) {
    const index = fileList.value.indexOf(file)
    const newFileList = fileList.value.slice()
    newFileList.splice(index, 1)
    fileList.value = newFileList
  }
}
</script>

<style lang="less"></style>
