/**
 * 生成儿童AI玩具产品方案分析文档
 */

import { handleWordTool } from './tools/word/index.js'

async function generateProductReport() {
  console.log('🚀 开始生成产品方案文档...\n')

  // 创建文档（使用软件文档主题）
  const createResult = await handleWordTool('word_create', {
    title: '儿童AI玩具产品方案分析',
    theme: 'software',
    author: '项目调研组',
    outputPath: '/Users/yangjiahui/Desktop/儿童AI玩具产品方案分析.docx'
  })
  const docId = JSON.parse(createResult.content[0].text).docId
  console.log('✅ 文档创建成功:', docId)

  // 封面信息
  await handleWordTool('word_add_heading', { docId, text: '儿童AI玩具产品方案分析', level: 1 })
  await handleWordTool('word_add_paragraph', { docId, text: '报告日期：2026年1月', bold: true })
  await handleWordTool('word_add_paragraph', { docId, text: '编制单位：项目调研组', bold: true })

  // ========== 一、项目背景 ==========
  await handleWordTool('word_add_heading', { docId, text: '一、项目背景', level: 2 })

  await handleWordTool('word_add_heading', { docId, text: '1.1 项目定位', level: 3 })
  await handleWordTool('word_add_paragraph', {
    docId,
    text: '开发一款面向3-12岁儿童的AI智能陪伴玩具，主打英语学习陪伴垂直场景，采用开源技术方案进行二次开发，实现快速上市和成本可控。'
  })

  await handleWordTool('word_add_heading', { docId, text: '1.2 目标市场', level: 3 })
  await handleWordTool('word_add_list', {
    docId,
    items: [
      '目标价位：180-399元（主流消费区间）',
      '目标用户：90后父母家庭、二三线城市',
      '目标毛利：≥100%（加价倍率≥2倍）'
    ],
    ordered: false
  })

  await handleWordTool('word_add_heading', { docId, text: '1.3 竞争策略', level: 3 })
  await handleWordTool('word_add_table', {
    docId,
    headers: ['维度', '竞品现状', '我方策略'],
    rows: [
      ['定价', '300-500元为主', '180元切入，性价比突围'],
      ['功能', '大而全', '聚焦英语学习单一场景'],
      ['技术', '自研为主', '开源二次开发，快速迭代'],
      ['渠道', '全渠道铺开', '先电商验证，再拓展']
    ],
    style: 'striped'
  })

  // ========== 二、技术方案 ==========
  await handleWordTool('word_add_heading', { docId, text: '二、技术方案', level: 2 })

  await handleWordTool('word_add_heading', { docId, text: '2.1 开源方案选型', level: 3 })
  await handleWordTool('word_add_paragraph', { docId, text: '选定方案：xiaozhi-esp32', bold: true })
  await handleWordTool('word_add_table', {
    docId,
    headers: ['评估维度', '详情'],
    rows: [
      ['GitHub Stars', '23,200+'],
      ['开源协议', 'MIT License（商用无限制）'],
      ['社区活跃度', '高（持续更新维护）'],
      ['技术栈', 'ESP-IDF + WebSocket/MQTT'],
      ['文档完善度', '中文文档齐全']
    ],
    style: 'striped'
  })

  await handleWordTool('word_add_paragraph', { docId, text: '技术架构：', bold: true })

  // 使用 Mermaid 绘制架构图 - 优化布局
  const architectureDiagram = `
flowchart LR
    subgraph Input["输入"]
        MIC["🎤 麦克风<br/>INMP441"]
    end

    subgraph MCU["ESP32-S3 主控"]
        direction TB
        AUDIO["音频处理"]
        WIFI["WiFi"]
    end

    subgraph Cloud["云端服务"]
        direction TB
        ASR["语音识别"]
        LLM["大语言模型"]
        TTS["语音合成"]
    end

    subgraph Output["输出"]
        AMP["🔊 功放+扬声器<br/>MAX98357A"]
    end

    MIC --> AUDIO
    AUDIO <--> WIFI
    WIFI <-->|WebSocket| ASR
    ASR --> LLM
    LLM --> TTS
    TTS -->|音频流| WIFI
    AUDIO --> AMP
`

  await handleWordTool('word_add_diagram', {
    docId,
    mermaid: architectureDiagram,
    caption: '图1 AI玩具系统架构图',
    width: 15,
    theme: 'professional'
  })

  await handleWordTool('word_add_heading', { docId, text: '2.2 硬件选型', level: 3 })
  await handleWordTool('word_add_table', {
    docId,
    headers: ['组件', '型号', '单价（元）', '说明'],
    rows: [
      ['主控芯片', 'ESP32-S3 N16R8', '18.00', '16MB Flash + 8MB PSRAM'],
      ['麦克风', 'INMP441', '3.50', 'I2S数字输出，高灵敏度'],
      ['功放模块', 'MAX98357A', '2.50', 'I2S输入，3W输出'],
      ['扬声器', '4Ω 3W喇叭', '2.00', '腔体设计影响音质'],
      ['电池', '18650锂电池', '8.00', '2000mAh，续航4-6小时'],
      ['充电模块', 'TP4056', '1.50', 'Type-C接口'],
      ['电源管理', 'AMS1117-3.3', '0.50', '3.3V稳压'],
      ['PCB板', '定制PCB', '5.00', '含SMT贴片'],
      ['按键+LED', '组合件', '2.00', '开关+指示灯'],
      ['电子件小计', '-', '43.00', '批量采购价']
    ],
    style: 'striped'
  })

  await handleWordTool('word_add_heading', { docId, text: '2.3 接线方案', level: 3 })
  await handleWordTool('word_add_paragraph', { docId, text: 'INMP441 麦克风接线：', bold: true })
  await handleWordTool('word_add_table', {
    docId,
    headers: ['INMP441引脚', 'ESP32-S3引脚', '说明'],
    rows: [
      ['VDD', '3.3V', '电源正极'],
      ['GND', 'GND', '电源负极'],
      ['SD', 'GPIO 4', '数据输出'],
      ['WS', 'GPIO 5', '字选择'],
      ['SCK', 'GPIO 6', '时钟'],
      ['L/R', 'GND', '左声道模式']
    ],
    style: 'professional'
  })

  await handleWordTool('word_add_paragraph', { docId, text: 'MAX98357A 功放接线：', bold: true })
  await handleWordTool('word_add_table', {
    docId,
    headers: ['MAX98357A引脚', 'ESP32-S3引脚', '说明'],
    rows: [
      ['VIN', '5V', '电源正极'],
      ['GND', 'GND', '电源负极'],
      ['DIN', 'GPIO 16', '数据输入'],
      ['BCLK', 'GPIO 15', '位时钟'],
      ['LRC', 'GPIO 7', '左右声道时钟']
    ],
    style: 'professional'
  })

  // ========== 三、成本核算 ==========
  await handleWordTool('word_add_heading', { docId, text: '三、成本核算', level: 2 })

  await handleWordTool('word_add_heading', { docId, text: '3.1 BOM成本明细', level: 3 })
  await handleWordTool('word_add_paragraph', { docId, text: '180元定价版本（基础款）：', bold: true })
  await handleWordTool('word_add_table', {
    docId,
    headers: ['类别', '项目', '成本（元）'],
    rows: [
      ['电子元器件', '主控+音频模块+电源', '43.00'],
      ['外壳', '注塑外壳（简约造型）', '8.00'],
      ['毛绒', '可拆卸毛绒外套', '12.00'],
      ['包装', '彩盒+说明书+配件', '5.00'],
      ['组装', '人工+测试', '8.00'],
      ['物流', '工厂到仓', '2.00'],
      ['硬件总成本', '-', '78.00']
    ],
    style: 'striped'
  })

  await handleWordTool('word_add_paragraph', { docId, text: '云服务成本（按月均摊）：', bold: true })
  await handleWordTool('word_add_table', {
    docId,
    headers: ['项目', '月成本/台', '说明'],
    rows: [
      ['ASR服务', '0.50', '按调用量计费'],
      ['LLM服务', '1.50', '国产大模型'],
      ['TTS服务', '0.50', '语音合成'],
      ['服务器', '0.30', '转发服务'],
      ['月均', '2.80', '按活跃用户计']
    ],
    style: 'striped'
  })

  await handleWordTool('word_add_heading', { docId, text: '3.2 利润测算', level: 3 })
  await handleWordTool('word_add_paragraph', { docId, text: '单品利润计算（180元定价）：', bold: true })
  await handleWordTool('word_add_table', {
    docId,
    headers: ['项目', '金额（元）', '占比'],
    rows: [
      ['零售价', '180.00', '100%'],
      ['硬件成本', '78.00', '43.3%'],
      ['毛利', '102.00', '56.7%'],
      ['加价倍率', '2.31倍', '>100%']
    ],
    style: 'professional'
  })

  await handleWordTool('word_add_paragraph', { docId, text: '不同定价毛利对比：', bold: true })
  await handleWordTool('word_add_table', {
    docId,
    headers: ['定价（元）', '硬件成本', '毛利', '毛利率', '加价倍率'],
    rows: [
      ['180', '78', '102', '56.7%', '2.31x'],
      ['249', '78', '171', '68.7%', '3.19x'],
      ['299', '85', '214', '71.6%', '3.52x'],
      ['399', '95', '304', '76.2%', '4.20x']
    ],
    style: 'striped'
  })
  await handleWordTool('word_add_paragraph', { docId, text: '注：299/399元版本含更好外壳和包装', style: 'note' })

  await handleWordTool('word_add_heading', { docId, text: '3.3 盈亏平衡分析', level: 3 })
  await handleWordTool('word_add_paragraph', { docId, text: '首批生产1000台：', bold: true })
  await handleWordTool('word_add_table', {
    docId,
    headers: ['项目', '金额（元）'],
    rows: [
      ['模具费用', '15,000'],
      ['首批物料', '78,000'],
      ['包装印刷', '5,000'],
      ['测试设备', '3,000'],
      ['总投入', '101,000']
    ],
    style: 'professional'
  })

  await handleWordTool('word_add_table', {
    docId,
    headers: ['定价', '毛利/台', '盈亏平衡点'],
    rows: [
      ['180元', '102元', '990台'],
      ['249元', '171元', '591台'],
      ['299元', '214元', '472台']
    ],
    style: 'striped'
  })

  // ========== 四、产品规划 ==========
  await handleWordTool('word_add_heading', { docId, text: '四、产品规划', level: 2 })

  await handleWordTool('word_add_heading', { docId, text: '4.1 产品矩阵', level: 3 })
  await handleWordTool('word_add_table', {
    docId,
    headers: ['版本', '定价', '定位', '特点'],
    rows: [
      ['基础款', '180元', '引流款', '核心功能，简约外观'],
      ['标准款', '299元', '主力款', '优化音质，精致外观'],
      ['IP联名款', '399元', '利润款', '热门IP授权，限量发售']
    ],
    style: 'striped'
  })

  await handleWordTool('word_add_heading', { docId, text: '4.2 功能规划', level: 3 })
  await handleWordTool('word_add_paragraph', { docId, text: 'MVP版本（V1.0）核心功能：', bold: true })
  await handleWordTool('word_add_list', {
    docId,
    items: [
      '英语对话练习 - 日常口语、场景对话',
      '单词学习 - 跟读、拼写、造句',
      '睡前故事 - 英文绘本故事朗读',
      '情感陪伴 - 基础情感交互'
    ],
    ordered: true
  })

  await handleWordTool('word_add_paragraph', { docId, text: '后续迭代（V2.0）：', bold: true })
  await handleWordTool('word_add_list', {
    docId,
    items: [
      '学习报告推送（家长端）',
      '个性化学习路径',
      '多角色扮演对话',
      '会员订阅增值服务'
    ],
    ordered: true
  })

  await handleWordTool('word_add_heading', { docId, text: '4.3 差异化卖点', level: 3 })
  await handleWordTool('word_add_table', {
    docId,
    headers: ['卖点', '描述', '竞品对比'],
    rows: [
      ['专注英语', '垂直场景深耕', '竞品功能大而全'],
      ['纯正发音', '美式/英式可选', '部分竞品发音机械'],
      ['自适应难度', '根据水平调整', '多数竞品固定难度'],
      ['家长可控', '时长/内容管控', '部分竞品缺失']
    ],
    style: 'striped'
  })

  // ========== 五、开发计划 ==========
  await handleWordTool('word_add_heading', { docId, text: '五、开发计划', level: 2 })

  await handleWordTool('word_add_heading', { docId, text: '5.1 项目里程碑', level: 3 })
  await handleWordTool('word_add_table', {
    docId,
    headers: ['阶段', '周期', '交付物'],
    rows: [
      ['技术验证', '2周', '开发板功能演示'],
      ['原型开发', '4周', '工程样机3台'],
      ['小批试产', '3周', '100台测试机'],
      ['用户测试', '2周', '测试报告+优化清单'],
      ['量产准备', '2周', '量产BOM+生产SOP'],
      ['首批量产', '3周', '1000台成品'],
      ['总计', '16周', '约4个月']
    ],
    style: 'striped'
  })

  await handleWordTool('word_add_heading', { docId, text: '5.2 团队配置', level: 3 })
  await handleWordTool('word_add_table', {
    docId,
    headers: ['角色', '人数', '职责'],
    rows: [
      ['项目经理', '1', '整体协调、进度把控'],
      ['嵌入式工程师', '1', '固件开发、硬件调试'],
      ['后端工程师', '1', '云服务对接、API开发'],
      ['产品经理', '1', '需求定义、用户测试'],
      ['工业设计', '外包', '外观设计、结构设计'],
      ['供应链', '1', '采购、生产对接']
    ],
    style: 'striped'
  })

  await handleWordTool('word_add_heading', { docId, text: '5.3 风险控制', level: 3 })
  await handleWordTool('word_add_table', {
    docId,
    headers: ['风险', '等级', '应对措施'],
    rows: [
      ['技术风险', '中', '基于成熟开源方案，降低不确定性'],
      ['供应链风险', '低', '珠三角供应链成熟，备选供应商'],
      ['市场风险', '中', '小批量验证PMF后再规模化'],
      ['合规风险', '中', '提前申请3C认证、儿童产品标准']
    ],
    style: 'professional'
  })

  // ========== 六、商业模式 ==========
  await handleWordTool('word_add_heading', { docId, text: '六、商业模式', level: 2 })

  await handleWordTool('word_add_heading', { docId, text: '6.1 收入结构', level: 3 })
  await handleWordTool('word_add_paragraph', { docId, text: '收入来源构成（预期）：' })
  await handleWordTool('word_add_list', {
    docId,
    items: [
      '硬件销售：70%',
      '会员订阅：25%',
      '增值服务：5%'
    ],
    ordered: false
  })

  await handleWordTool('word_add_heading', { docId, text: '6.2 会员体系设计', level: 3 })
  await handleWordTool('word_add_table', {
    docId,
    headers: ['等级', '月费', '权益'],
    rows: [
      ['免费版', '0元', '基础对话（50次/天）'],
      ['标准会员', '19.9元/月', '无限对话+学习报告'],
      ['高级会员', '39.9元/月', '定制课程+1对1辅导']
    ],
    style: 'striped'
  })

  await handleWordTool('word_add_paragraph', { docId, text: 'LTV测算：', bold: true })
  await handleWordTool('word_add_list', {
    docId,
    items: [
      '假设30%用户转化为付费会员',
      '平均会员周期6个月',
      '单用户LTV = 180 + 19.9×6×0.3 = 215.8元'
    ],
    ordered: false
  })

  await handleWordTool('word_add_heading', { docId, text: '6.3 渠道策略', level: 3 })
  await handleWordTool('word_add_paragraph', { docId, text: '第一阶段：电商验证', bold: true })
  await handleWordTool('word_add_list', {
    docId,
    items: [
      '抖音直播：测试转化率、收集反馈',
      '拼多多：下沉市场试水',
      '淘宝/京东：品牌店铺建设'
    ],
    ordered: false
  })

  await handleWordTool('word_add_paragraph', { docId, text: '第二阶段：渠道拓展', bold: true })
  await handleWordTool('word_add_list', {
    docId,
    items: [
      '母婴连锁：孩子王、爱婴室',
      '教育机构：英语培训机构合作',
      '礼品渠道：企业定制、节日礼盒'
    ],
    ordered: false
  })

  // ========== 七、财务预测 ==========
  await handleWordTool('word_add_heading', { docId, text: '七、财务预测', level: 2 })

  await handleWordTool('word_add_heading', { docId, text: '7.1 销量预测（首年）', level: 3 })
  await handleWordTool('word_add_table', {
    docId,
    headers: ['季度', '销量（台）', '说明'],
    rows: [
      ['Q1', '2,000', '小批量测试+种子用户'],
      ['Q2', '5,000', '电商渠道起量'],
      ['Q3', '10,000', '扩展新渠道'],
      ['Q4', '15,000', '双11/双12爆发'],
      ['全年', '32,000', '']
    ],
    style: 'striped'
  })

  await handleWordTool('word_add_heading', { docId, text: '7.2 收入预测（首年）', level: 3 })
  await handleWordTool('word_add_table', {
    docId,
    headers: ['项目', '金额（万元）', '说明'],
    rows: [
      ['硬件销售', '576', '均价180元×32000台'],
      ['会员订阅', '69', '30%转化×6个月×19.9元'],
      ['总收入', '645', '']
    ],
    style: 'professional'
  })

  await handleWordTool('word_add_heading', { docId, text: '7.3 成本预测（首年）', level: 3 })
  await handleWordTool('word_add_table', {
    docId,
    headers: ['项目', '金额（万元）', '说明'],
    rows: [
      ['硬件成本', '250', '78元×32000台'],
      ['云服务', '27', '2.8元×32000台×3月'],
      ['研发人力', '80', '4人×20万/年'],
      ['营销费用', '100', '15%营销占比'],
      ['运营成本', '30', '仓储物流客服'],
      ['总成本', '487', '']
    ],
    style: 'striped'
  })

  await handleWordTool('word_add_heading', { docId, text: '7.4 盈利预测', level: 3 })
  await handleWordTool('word_add_table', {
    docId,
    headers: ['项目', '金额（万元）'],
    rows: [
      ['总收入', '645'],
      ['总成本', '487'],
      ['净利润', '158'],
      ['净利率', '24.5%']
    ],
    style: 'professional'
  })

  // ========== 八、结论与建议 ==========
  await handleWordTool('word_add_heading', { docId, text: '八、结论与建议', level: 2 })

  await handleWordTool('word_add_heading', { docId, text: '8.1 可行性结论', level: 3 })
  await handleWordTool('word_add_list', {
    docId,
    items: [
      '技术可行：基于xiaozhi-esp32开源方案，技术成熟度高',
      '成本可控：硬件BOM成本78元，180元定价可实现130%加价倍率',
      '市场可期：行业CAGR>50%，市场空间充足',
      '风险可控：小批量验证模式，投入可控'
    ],
    ordered: true
  })

  await handleWordTool('word_add_heading', { docId, text: '8.2 执行建议', level: 3 })
  await handleWordTool('word_add_table', {
    docId,
    headers: ['优先级', '建议'],
    rows: [
      ['P0', '立即启动技术验证，2周内完成原型'],
      ['P0', '确定供应链合作伙伴'],
      ['P1', '申请相关资质认证（3C等）'],
      ['P1', '设计产品外观和包装'],
      ['P2', '规划电商渠道和营销策略']
    ],
    style: 'striped'
  })

  await handleWordTool('word_add_heading', { docId, text: '8.3 关键成功因素', level: 3 })
  await handleWordTool('word_add_list', {
    docId,
    items: [
      '产品体验：交互延迟<1秒，语音识别准确率>90%',
      '成本控制：BOM成本控制在80元以内',
      '快速迭代：基于用户反馈持续优化',
      '内容运营：持续更新优质英语学习内容'
    ],
    ordered: true
  })

  // ========== 附录 ==========
  await handleWordTool('word_add_heading', { docId, text: '附录：技术参考', level: 2 })

  await handleWordTool('word_add_heading', { docId, text: 'A. xiaozhi-esp32项目信息', level: 3 })
  await handleWordTool('word_add_list', {
    docId,
    items: [
      'GitHub地址：https://github.com/78/xiaozhi-esp32',
      '开源协议：MIT License',
      '技术文档：项目Wiki完整中文文档',
      '社区支持：活跃的开发者社区'
    ],
    ordered: false
  })

  await handleWordTool('word_add_heading', { docId, text: 'B. 推荐供应商', level: 3 })
  await handleWordTool('word_add_table', {
    docId,
    headers: ['类别', '推荐供应商', '备注'],
    rows: [
      ['芯片', '立创商城、华秋商城', '小批量首选'],
      ['PCB', '嘉立创、捷配', '打样快速'],
      ['注塑', '东莞模具厂', '成本低'],
      ['组装', '深圳PCBA厂', '一站式服务']
    ],
    style: 'striped'
  })

  await handleWordTool('word_add_heading', { docId, text: 'C. 认证要求', level: 3 })
  await handleWordTool('word_add_table', {
    docId,
    headers: ['认证', '必要性', '周期'],
    rows: [
      ['3C认证', '必须', '4-6周'],
      ['SRRC无线电型号核准', '必须', '3-4周'],
      ['GB 6675玩具安全', '建议', '2-3周']
    ],
    style: 'professional'
  })

  await handleWordTool('word_add_paragraph', { docId, text: '报告编制完成', bold: true })

  // 保存文档
  const saveResult = await handleWordTool('word_save', {
    docId,
    addPageNumbers: true,
    addHeader: '儿童AI玩具产品方案分析',
    addFooter: '内部资料 请勿外传'
  })

  console.log('\n✅ 文档生成完成!')
  console.log(JSON.parse(saveResult.content[0].text))
}

generateProductReport().catch(console.error)
