/**
 * 测试 Mermaid 图表功能
 */

import { handleWordTool } from './tools/word/index.js'

async function testDiagram() {
  console.log('🚀 测试 Mermaid 图表功能...\n')

  // 1. 创建文档
  const createResult = await handleWordTool('word_create', {
    title: 'Mermaid 图表测试',
    theme: 'software',
    outputPath: '/Users/yangjiahui/Desktop/Mermaid图表测试.docx'
  })
  const docId = JSON.parse(createResult.content[0].text).docId
  console.log('✅ 文档创建成功:', docId)

  // 2. 添加标题
  await handleWordTool('word_add_heading', {
    docId,
    text: 'Mermaid 图表功能演示',
    level: 1
  })

  // 3. 添加架构图（flowchart）
  await handleWordTool('word_add_heading', { docId, text: '一、系统架构图', level: 2 })
  await handleWordTool('word_add_paragraph', {
    docId,
    text: '以下是 AI 玩具的系统架构图，使用 Mermaid flowchart 语法绘制：'
  })

  const architectureDiagram = `
flowchart TB
    subgraph Cloud["云端服务"]
        ASR["ASR<br/>语音识别"]
        LLM["LLM<br/>大语言模型"]
        TTS["TTS<br/>语音合成"]
    end

    subgraph Device["硬件终端"]
        ESP32["ESP32-S3<br/>主控芯片"]
        MIC["INMP441<br/>麦克风"]
        SPK["MAX98357A<br/>功放+扬声器"]
    end

    MIC -->|音频采集| ESP32
    ESP32 -->|WebSocket| ASR
    ASR -->|文本| LLM
    LLM -->|回复| TTS
    TTS -->|音频流| ESP32
    ESP32 -->|播放| SPK
`

  const result1 = await handleWordTool('word_add_diagram', {
    docId,
    mermaid: architectureDiagram,
    caption: '图1 AI玩具系统架构图',
    width: 14,
    theme: 'default'
  })
  console.log('架构图:', JSON.parse(result1.content[0].text).message)

  // 4. 添加流程图
  await handleWordTool('word_add_heading', { docId, text: '二、交互流程图', level: 2 })
  await handleWordTool('word_add_paragraph', {
    docId,
    text: '用户与 AI 玩具的交互流程：'
  })

  const flowDiagram = `
flowchart LR
    A[用户说话] --> B[麦克风采集]
    B --> C[语音识别]
    C --> D[AI理解&生成]
    D --> E[语音合成]
    E --> F[扬声器播放]
    F --> G[用户听到回复]
`

  const result2 = await handleWordTool('word_add_diagram', {
    docId,
    mermaid: flowDiagram,
    caption: '图2 语音交互流程',
    width: 15
  })
  console.log('流程图:', JSON.parse(result2.content[0].text).message)

  // 5. 添加时序图
  await handleWordTool('word_add_heading', { docId, text: '三、时序图', level: 2 })
  await handleWordTool('word_add_paragraph', {
    docId,
    text: '系统各模块的调用时序：'
  })

  const sequenceDiagram = `
sequenceDiagram
    participant U as 用户
    participant D as 设备
    participant S as 云服务

    U->>D: 说话
    D->>S: 发送音频
    S->>S: ASR识别
    S->>S: LLM处理
    S->>S: TTS合成
    S->>D: 返回音频
    D->>U: 播放回复
`

  const result3 = await handleWordTool('word_add_diagram', {
    docId,
    mermaid: sequenceDiagram,
    caption: '图3 系统调用时序图',
    width: 12
  })
  console.log('时序图:', JSON.parse(result3.content[0].text).message)

  // 6. 添加状态图
  await handleWordTool('word_add_heading', { docId, text: '四、状态图', level: 2 })
  await handleWordTool('word_add_paragraph', {
    docId,
    text: '设备的运行状态：'
  })

  const stateDiagram = `
stateDiagram-v2
    [*] --> Idle: 开机
    Idle --> Listening: 唤醒词
    Listening --> Processing: 说完话
    Processing --> Speaking: 收到回复
    Speaking --> Idle: 播放完成
    Idle --> Sleep: 超时
    Sleep --> Idle: 唤醒
    Sleep --> [*]: 关机
`

  const result4 = await handleWordTool('word_add_diagram', {
    docId,
    mermaid: stateDiagram,
    caption: '图4 设备状态转换图',
    width: 12
  })
  console.log('状态图:', JSON.parse(result4.content[0].text).message)

  // 7. 保存文档
  const saveResult = await handleWordTool('word_save', {
    docId,
    addPageNumbers: true,
    addHeader: 'Mermaid 图表测试',
    addFooter: '测试文档'
  })

  console.log('\n✅ 测试完成!')
  console.log(JSON.parse(saveResult.content[0].text))
}

testDiagram().catch(console.error)
