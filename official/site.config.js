/* ============================================================================
   站点内容配置 —— 素材补齐入口
   ----------------------------------------------------------------------------
   所有文案、图片、指标都集中在这里。任何一项留 null / 空数组，页面框架照常
   显示，对应位置自动降级为：
     · 图片   → 骨架屏（#069，0.9s 扫光）
     · 列表   → 空状态（#072，0.35s 淡入，线描图标 + 「待补充」+ 说明）
   后续补素材只改这个文件，不用动 HTML / CSS / JS。

   约定：
     - 图片路径相对 _preview/ 目录；也可填完整 URL
     - 填 null 表示「尚未提供」，框架会保留位置和比例，不会塌陷
     - boards 是已经做好的数据页面，路径按实际部署位置改
   ============================================================================ */
window.SITE = {
  /* 品牌 —— 未提供的用 null */
  brand: {
    name: "顶呱呱杭州分公司",
    shortName: "顶呱呱杭州分公司",
    logo: "assets/dgg-logo.png",   // 官网公开 Logo（http://www.cddgg.com）
    slogan: "受尊重的服务者，让老板经营企业更简单"
  },

  /* 顶部公告条 */
  announcement: "一站式企业服务就找顶呱呱 · 全国咨询热线 400 666 5987",

  /* 主导航（巨型菜单 #067 的分组内容） */
  nav: [
    { label: "关于我们", href: "#about", mega: [
      { title: "公司", links: [["企业介绍", "#about"], ["发展历程", "#timeline"], ["联系我们", "#cta"]] },
      { title: "业务", links: [["工商服务", "#solutions"], ["财税服务", "#solutions"], ["法律服务", "#solutions"]] },
      { title: "更多", links: [["资质许可", "#solutions"], ["知识产权", "#solutions"], ["互联网服务", "#solutions"]] }
    ]},
    { label: "业务与服务", href: "#solutions", mega: null },
    { label: "发展历程", href: "#timeline", mega: null },
    { label: "公司团队", href: "#team", mega: null },
    { label: "新闻洞察", href: "#news", mega: null }
  ],

  /* Hero ---------- 内容来自官网公开描述 */
  hero: {
    eyebrow: "杭州分公司 - 一站式企业服务平台",
    title: "让老板经营企业更简单",
    highlight: "更简单",
    subtitle: "扎根杭州拱墅，为本地及长三角企业提供工商注册、代理记账、知识产权、法律服务、资质许可与互联网服务，全程专属顾问一对一跟进。",
    primaryCta: "了解公司",
    secondaryCta: "业务与服务",
    image: "assets/hero-poster.jpg",  // 视频封面图；视频缺失时显示它，框架不塌
    bgImage: null,
    /* 置顶短片路径写在 videoChannel.pinnedVideo；此处留 null 走同一份配置，
       避免两处路径不一致。留空则显示封面 + 官方观看入口，不 404、不白屏。 */
    video: null
  },

  /* 视频号信息（顶呱呱-企业服务）—— 取自官方分享页公开元数据 */
  videoChannel: {
    name: "顶呱呱-企业服务",
    desc: "集团官方视频号 · 企业服务宣传片",
    avatar: "assets/video-avatar.jpg",
    cover: "assets/video-cover.jpg",
    title: "60s带您走进顶呱呱！",
    tags: "#顶呱呱集团  #企业服务  #宣传片  #老板  #创业",
    date: "2025-09-11",
    location: "成都市",
    stats: { fav: "28", forward: "602", like: "29", comment: "5" },
    url: "https://weixin.qq.com/sph/AtfmPHvJva",
    poster: "assets/hero-poster.jpg",
    /* 用户确认：不显示时长。留空字符串；以后要显示时填 "01:00" 即可 */
    duration: "",
    /* 置顶短片：微信里导出的原始 mp4，52.83s / 1280x720 / 9.2MB。
       Hero 会自动 <video autoplay loop muted playsinline> 循环播放；
       设为 null 则退回封面 + 官方观看入口。 */
    pinnedVideo: "assets/pinned.mp4"
  },

  /* 数据指标条 —— 集团公开口径，需确认是否适用于杭州分公司 */
  metrics: [
    { value: "2017", unit: "", label: "杭州分公司成立" },
    { value: "9", unit: "大", label: "业务业态" },
    { value: "5000", unit: "+", label: "专业顾问" },
    { value: "530", unit: "万+", label: "累计服务企业" }
  ],

  /* 解决方案 —— 卡片正面用官网蓝色业务图标；点开后的详情用官网实景照 + 业务介绍 */
  solutions: [
    {
      icon: "▦", title: "工商服务",
      short: "注册 · 变更 · 注销",
      desc: "公司注册、公司变更、企业注销等全生命周期工商办理。",
      image: "assets/sol-1.png",
      detail: {
        qr: null,                    // 该业务专属二维码；留 null 即统一使用 contact.qr
        image: "assets/svc-1.png",
        intro: "覆盖企业设立到注销的全流程工商代办。专业顾问一对一跟进，流程透明、合同明码标价、无隐形费用，让创业者少花钱、少跑路。",
        items: ["有限公司 / 个体工商户 / 集团公司注册", "公司变更、收购与注销", "银行开户、印章与执照补办", "工商咨询与企业核名"]
      }
    },
    {
      icon: "◫", title: "财税服务",
      short: "记账 · 报税 · 筹划",
      desc: "代理记账、税务筹划、财务规范与经营数据支持。",
      image: "assets/sol-2.png",
      detail: {
        qr: null,
        image: "assets/svc-2.png",
        intro: "面向中小微企业的代理记账、税务筹划与财务规范服务。集团持有 AAAAA 代理记账许可证，服务流程标准化，账务处理与申报节点全程可查。",
        items: ["代理记账与纳税申报", "税务筹划与财务规范", "财务外包与经营数据支持", "年度汇算清缴与工商年报"]
      }
    },
    {
      icon: "◇", title: "法律服务",
      short: "合同 · 合规 · 咨询",
      desc: "合同审查、法律咨询与企业合规支持。",
      image: "assets/sol-3.png",
      detail: {
        qr: null,
        image: "assets/svc-3.png",
        intro: "为企业经营提供合同审查、法务合规与法律咨询支持，覆盖设立、经营、扩张到退出各阶段的法律风险防控，由专业顾问团队响应。",
        items: ["合同起草与审查", "企业合规与风控体系", "劳动用工法律咨询", "股权结构与纠纷咨询"]
      }
    },
    {
      icon: "◈", title: "知识产权",
      short: "商标 · 版权 · 保护",
      desc: "商标注册、版权登记与知识产权保护。",
      image: "assets/sol-4.png",
      detail: {
        qr: null,
        image: "assets/svc-4.png",
        intro: "集团持有备案商标代理资质，提供商标注册、版权登记、专利申请与知识产权保护的一站式服务，帮助企业把创新成果转化为可运营的资产。",
        items: ["商标注册、续展与异议", "版权 / 软件著作权登记", "专利申请与年费维护", "知识产权监测与维权"]
      }
    },
    {
      icon: "▤", title: "资质许可",
      short: "ICP · 食品 · 医疗",
      desc: "各类经营资质与行政许可代办服务。",
      image: "assets/sol-5.jpg",
      detail: {
        qr: null,
        image: "assets/svc-5.jpg",
        intro: "各类经营资质与行政许可代办，覆盖 ICP 许可证、食品经营许可、医疗器械经营许可、卫生许可与互联网资质等，渠道正规、办理手续合规。",
        items: ["ICP / 互联网信息服务资质", "食品经营、卫生许可", "医疗器械经营许可", "建筑与行业专项资质"]
      }
    },
    {
      icon: "◎", title: "互联网服务",
      short: "网站 · 小程序 · 营销",
      desc: "网站建设、小程序与数字化营销支持。",
      image: "assets/sol-6.jpg",
      detail: {
        qr: null,
        image: "assets/svc-6.jpg",
        intro: "依托集团自主研发的智能企服系统与多项软件著作权，为企业提供官网建设、小程序开发与数字化营销支持，用数字化赋能服务升级。",
        items: ["官网 / 品牌落地页建设", "微信小程序开发", "数字化营销与获客", "智能企服系统对接"]
      }
    }
  ],

  /* 关于我们 */
  about: {
    title: "关于顶呱呱杭州分公司",
    body: "顶呱呱杭州分公司落地拱墅，是集团布局华东市场的核心阵地。团队为杭州及长三角企业提供工商注册、代理记账、知识产权、法律服务、资质许可与互联网服务，从创业筹备到成熟经营，全程由专属顾问一对一跟进，合同明码标价、服务流程可查。",
    image: "assets/office-front.jpg",
    points: ["扎根杭州拱墅 · 华东核心阵地", "工商财税一站式办理", "专属顾问一对一跟进", "服务流程透明可查"],
    address: "拱墅区绍兴路 88 号兴和元洲园区 2 号楼 5 层",
    addressNote: "顶呱呱杭州分公司"
  },

  /* 发展历程（滚动时间轴 #145）—— 公开资料整理，待核对 */
  timeline: [
    { year: "2008", title: "集团成立", desc: "顶呱呱在四川成都成立，开启一站式企业服务。" },
    { year: "2016", title: "科技化升级", desc: "顶呱呱科技股份有限公司成立，推进智能企服平台建设。" },
    { year: "2017", title: "落地杭州", desc: "杭州分公司在拱墅成立，成为集团布局华东市场的核心阵地。" },
    { year: "至今", title: "深耕华东", desc: "服务杭州及长三角企业，覆盖工商、财税、知识产权、法律、资质与互联网六大业务线。" }
  ],

  /* 发展历程右侧 —— 素材与文案均取自官网「关于我们 /gywm」的资质与承诺 */
  timelineAside: {
    image: "assets/history-side.jpg",
    kicker: "Qualifications",
    title: "以标准化服务，赢得长期信任",
    intro: "集团资质齐全，服务流程标准化、透明化，合同明码标价，无隐形费用。",
    items: [
      "持有备案商标代理资质",
      "AAAAA 代理记账许可证",
      "省市认定中小企业公共服务示范平台",
      "国家服务业发展引导资金支持单位",
      "ISO9001 质量体系认证",
      "中国最佳企业服务商"
    ],
    stats: [
      { value: "2008", label: "集团始创" },
      { value: "15", label: "核心城市" },
      { value: "530", label: "万家企业" }
    ],
    note: "六大品质服务承诺 · 7 天不满意退款，超时与超期均有责必赔。"
  },

  /* 企业文化 —— 文本取自官网「关于我们 /gywm」：使命 / 愿景 / 价值观四条解读 / 服务优势六项 */
  culture: {
    core: [
      { no: "01", en: "Mission", title: "使命", desc: "受尊重的服务者，让老板经营企业更简单。" },
      { no: "02", en: "Vision", title: "愿景", desc: "让更多人，生活更美好。" },
      { no: "03", en: "Values", title: "价值观", desc: "以客户为中心 · 何为正确 · 合作共赢 · 全力以赴。" }
    ],
    values: [
      { no: "01", title: "对客户好", desc: "把客户的事当成自己的事，先解决问题，再谈其它。" },
      { no: "02", title: "诚实守信", desc: "承诺必达；不尽责即无诚信，对客户、同事与公司都算数。" },
      { no: "03", title: "合作共赢", desc: "不设部门墙，一起把事情做成，成果与伙伴共享。" },
      { no: "04", title: "全力以赴", desc: "把岗位做到专家级，让自己成为团队不可或缺的一员。" }
    ],
    advantage: [
      { icon: "flow", title: "全程", desc: "服务全流程可视化" },
      { icon: "bolt", title: "高效", desc: "专属顾问一对一沟通" },
      { icon: "target", title: "精准", desc: "拥有各类商业查询工具" },
      { icon: "glass", title: "透明", desc: "服务产品价格透明" },
      { icon: "shield", title: "安全", desc: "多重监督，售后保障" },
      { icon: "grid", title: "便捷", desc: "提供一站式企业服务" }
    ]
  },

  /* 公司影像（轮播 #058）—— 素材来自「顶呱呱集团」公众号 */
  cases: [
    { name: "品牌前台", desc: "政企一站式服务专家 · 杭州新办公场地", image: "assets/office-front.jpg" },
    { name: "洽谈空间", desc: "新中式景观与多元商务洽谈区", image: "assets/office-lounge.jpg" },
    { name: "会议空间", desc: "全景落地窗会议室与智能投屏设备", image: "assets/office-meeting.jpg" },
    { name: "会客区域", desc: "兼顾专业质感与人文温度的接待空间", image: "assets/office-talks.jpg" }
  ],

  /* 新闻洞察 —— 公众号「顶呱呱集团」文章，杭州分公司相关全部置顶 */
  news: [
    {
      date: "2026-07-07", tag: "杭州分公司", pinned: true,
      title: "七年深耕再启新程｜顶呱呱杭州分公司乔迁盛典圆满礼成",
      desc: "7 月 6 日乔迁庆典隆重举行。杭州分公司自 2017 年落地以来，始终是集团布局华东市场的核心阵地。",
      image: "assets/event-qiaqian.jpg",
      href: "http://mp.weixin.qq.com/s?__biz=MzA5MDI5MTcxNQ==&mid=2650590817&idx=1&sn=44314e57d7b47ece7a0bbf22a1cf6187"
    },
    {
      date: "2026-06-22", tag: "杭州分公司", pinned: true,
      title: "焕新启程，服务升级｜顶呱呱杭州公司新办公场地正式启用",
      desc: "全新办公场地投入运营，整合工商财税、资质代办、软件开发、跨境注册等全品类业务资源。",
      image: "assets/office-lounge.jpg",
      href: "http://mp.weixin.qq.com/s?__biz=MzA5MDI5MTcxNQ==&mid=2650590733&idx=2&sn=3e85c712048c5c69e8d9a71f7f6c34b6"
    },
    {
      date: "2025-12-19", tag: "杭州分公司", pinned: true,
      title: "政企协同，聚力赋能丨杭州市经信局领导一行莅临顶呱呱集团参观交流",
      desc: "围绕政策精准触达、专业服务集成、数字化赋能与服务生态共建，共探中小企业服务新体系。",
      image: "assets/gov-visit.jpg",
      href: "http://mp.weixin.qq.com/s?__biz=MzA5MDI5MTcxNQ==&mid=2650589959&idx=1&sn=2d2f7568d5c42df6b9300026f199a877"
    }
  ],

  /* ------------------------------------------------------------------
     公司人员介绍：7 位负责人（总经理 + 6 位事业部总/总监），点击可展开该部门成员。
     负责人大图 560×700 → assets/team/；成员小图 240×300 → assets/team/m/。
     职务与简介为占位文案，待 HR 确认后替换；秦瑞露暂无照片，留占位。
     ------------------------------------------------------------------ */
  team: {
    members: [
      { name: "秦瑞露", title: "杭州分公司总经理", dept: "杭州分公司", lead: true,
        photo: null,                     // 有照片后填 "assets/team/lead.jpg"
        desc: "企业服务行业深耕二十余年的老兵，从一线顾问一路做到分公司总经理，带出一支 47 人的专家团队，累计服务企业客户上万家，经营业绩常年稳居区域第一梯队。", people: [] },
      { name: "付辉", title: "事业部总经理", dept: "工商 · 会计事业部", img: "assets/team/gs1.jpg",
        desc: "从业十余年的工商财税双线老手，注册、变更、记账、注销的全链路烂熟于心，团队年交付数千单、差错率近乎为零，是客户口口相传的『靠得住』。",
        people: [
          { name: "于海岩", title: "工商注册主管", img: "assets/team/m/gs2.jpg", dept: "工商" },
          { name: "唐雷", title: "变更注销主管", img: "assets/team/m/gs3.jpg", dept: "工商" },
          { name: "吴桐青", title: "资深工商顾问", img: "assets/team/m/gs4.jpg", dept: "工商" },
          { name: "张玄琳", title: "工商顾问", img: "assets/team/m/gs5.jpg", dept: "工商" },
          { name: "王玉", title: "工商顾问", img: "assets/team/m/gs6.jpg", dept: "工商" },
          { name: "赫振祥", title: "资质联动顾问", img: "assets/team/m/gs7.jpg", dept: "工商" },
          { name: "吴甜", title: "代理记账主管", img: "assets/team/m/cw2.jpg", dept: "会计" },
          { name: "龙菲", title: "主办会计", img: "assets/team/m/cw3.jpg", dept: "会计" },
          { name: "龚莉", title: "财税顾问", img: "assets/team/m/cw4.jpg", dept: "会计" },
        ] },
      { name: "陈杰", title: "事业部总经理", dept: "法律事业部", img: "assets/team/fl1.jpg",
        desc: "法律行业摸爬滚打十余年，经手常年法律顾问与商事争议案件数百起，最擅长把复杂的法律问题讲成老板一听就懂的一二三，帮企业把风险拦在合同签订之前。",
        people: [
          { name: "庄原", title: "法务顾问", img: "assets/team/m/fl2.jpg" },
          { name: "张冉", title: "劳动人事顾问", img: "assets/team/m/fl3.jpg" },
          { name: "欧阳效晨", title: "知识产权顾问", img: "assets/team/m/fl4.jpg" },
          { name: "章雯", title: "公司治理顾问", img: "assets/team/m/fl5.jpg" },
          { name: "苗春旭", title: "诉讼支持顾问", img: "assets/team/m/fl6.jpg" },
          { name: "黄蜜军", title: "法律顾问", img: "assets/team/m/fl7.jpg" },
        ] },
      { name: "吴文杰", title: "事业部总经理", dept: "资质事业部", img: "assets/team/zz1.jpg",
        desc: "深耕资质许可领域十余年，各类行政许可、行业资质、体系认证的审批门道了然于胸，经手申报的通过率长期保持行业领先，难办的资质到他能找到办法。",
        people: [
          { name: "吕建波", title: "资质申报主管", img: "assets/team/m/zz2.jpg" },
          { name: "夏大雨", title: "审批对接主管", img: "assets/team/m/zz3.jpg" },
          { name: "吴可心", title: "许可顾问", img: "assets/team/m/zz4.jpg" },
          { name: "戴妙玲", title: "资质顾问", img: "assets/team/m/zz5.jpg" },
          { name: "吴雨", title: "体系认证顾问", img: "assets/team/m/zz6.jpg" },
          { name: "李徐博林", title: "项目申报顾问", img: "assets/team/m/zz7.jpg" },
          { name: "李鑫", title: "资质顾问", img: "assets/team/m/zz8.jpg" },
          { name: "赵珮瑶", title: "资质顾问", img: "assets/team/m/zz9.jpg" },
        ] },
      { name: "曾颖", title: "事业部总经理", dept: "跨境事业部", img: "assets/team/kj1.jpg",
        desc: "跨境服务老兵，十余年专注海外公司注册、年审与跨境合规，主流离岸属地均为一手渠道直达——别人跑一个月的事，他的团队一周就能交付。",
        people: [
          { name: "张云龙", title: "跨境注册顾问", img: "assets/team/m/kj2.jpg" },
          { name: "徐晨宇", title: "跨境合规顾问", img: "assets/team/m/kj3.jpg" },
          { name: "朱桂玲", title: "海外财税顾问", img: "assets/team/m/kj4.jpg" },
          { name: "胡岂铭", title: "跨境顾问", img: "assets/team/m/kj5.jpg" },
          { name: "石宇航", title: "跨境顾问", img: "assets/team/m/kj6.jpg" },
          { name: "秦辅辰", title: "跨境顾问", img: "assets/team/m/kj7.jpg" },
          { name: "胡晴颖", title: "跨境顾问", img: "assets/team/m/kj8.jpg" },
          { name: "陈柏羽", title: "跨境顾问", img: "assets/team/m/kj9.jpg" },
        ] },
      { name: "张一卫", title: "事业部总经理", dept: "互联网事业部", img: "assets/team/hl1.jpg",
        desc: "互联网与数字化服务十余年的实战派，ICP 备案、互联网资质、建站到数字化运营一条龙打通，项目上线速度稳居行业第一梯队，客户的官网与资质从不『卡壳』。",
        people: [
          { name: "何晓霞", title: "ICP 备案主管", img: "assets/team/m/hl2.jpg" },
          { name: "张也", title: "互联网资质顾问", img: "assets/team/m/hl3.jpg" },
          { name: "徐铁兵", title: "建站顾问", img: "assets/team/m/hl4.jpg" },
          { name: "李林杰", title: "数字化顾问", img: "assets/team/m/hl5.jpg" },
          { name: "王永志", title: "产品顾问", img: "assets/team/m/hl6.jpg" },
          { name: "舒田军", title: "运维顾问", img: "assets/team/m/hl7.jpg" },
          { name: "花东虎", title: "互联网顾问", img: "assets/team/m/hl8.jpg" },
          { name: "郝会超", title: "互联网顾问", img: "assets/team/m/hl9.jpg" },
          { name: "林贤华", title: "技术顾问", img: "assets/team/m/hl10.jpg" },
        ] },
      { name: "王梦洁", title: "营销中心总监", dept: "业务中心 · 营销中心", img: "assets/team/yx1.jpg",
        desc: "营销增长操盘手，十余年市场营销与客户成功经验，从获客、转化到复购搭建起完整的增长引擎，让分公司业绩连年保持两位数增长。",
        people: [
        ] },
    ]
  },

  /* 底部 CTA */
  cta: {
    title: "让企业经营更简单",
    desc: "有工商、财税、法律或资质方面的需求，欢迎随时联系我们。",
    button: "联系我们"
  },

  /* 咨询二维码 —— 全站「联系我们」统一使用下面这一张（巨型菜单、业务详情弹窗、
     联系弹窗共用）。换图只要替换 assets/qr/contact-wechat.png 即可；
     置为 null 时各处显示虚线占位框，版式不塌、不 404。 */
  contact: {
    qr: "assets/qr/contact-wechat.png",   // 微信二维码（联系我们统一使用）
    title: "微信咨询",
    desc: "扫码添加专属顾问，获取一对一服务方案",
    phone: "400 666 5987"
  },

  footer: {
    note: "© 顶呱呱杭州分公司 · 咨询热线 400 666 5987",
    address: "杭州分公司：杭州市拱墅区绍兴路 88 号兴和元洲园区 2 号楼 5 层",
    addressSub: "集团总部：四川省成都市武侯区科华北路 65 号世外桃源广场 A 座 3 楼",
    columns: []
  },

  /* ------------------------------------------------------------------
     员工通道：已经做好的数据页面（真实清单，不需要补素材）
     path 按最终部署位置调整；当前指向作品集仓库的 demo 目录
     ------------------------------------------------------------------ */

  /* ====== 员工通道顶部轮播（图片 + 文案，缺失图片时显示骨架） ====== */
  staffBanners: [
    { image: "assets/office-front.jpg", pos: "center",
      title: "顶呱呱杭州分公司", sub: "拱墅区绍兴路 88 号兴和元洲园区 2 号楼 5 层" },
    { image: "assets/boards/index.png", pos: "top center",
      title: "分公司数据看板", sub: "活动量 · 三力 · 业绩一屏总览，每日自动同步" },
    { image: "assets/boards/callqa.png", pos: "top center",
      title: "通话录音质检", sub: "827 条录音 · 五维评分 · 每周更新" },
    { image: "assets/office-meeting.jpg", pos: "center",
      title: "经营复盘会", sub: "数据驱动决策，每周一 09:30 准时开始" }
  ],

  /* ====== 员工通道：六大分区（取自看板导航页 nav.html）+ 公告 ====== */
  portal: [
  {
    "name": "业绩数据",
    "icon": "💰",
    "layout": "",
    "files": [
      {
        "icon": "📈",
        "title": "分公司业绩含续费",
        "short": "分公司业绩含续费",
        "desc": "业绩进度追踪 · 全月/周度/日度 · 业绩PK",
        "cover": "assets/boards/yj.png",
        "url": "../demo/yj.html",
        "tag": "在线看板"
      },
      {
        "icon": "🛒",
        "title": "分公司订单商品明细",
        "short": "分公司订单商品明细",
        "desc": "商品排名 · 分公司×业态热点图 · 多维筛选",
        "cover": "assets/boards/sku.png",
        "url": "../demo/sku.html",
        "tag": "在线看板"
      }
    ]
  },
  {
    "name": "集团看板",
    "icon": "🏢",
    "layout": "",
    "files": [
      {
        "icon": "🏃",
        "title": "集团活动量数据看板",
        "short": "集团活动量",
        "desc": "集团口径 · 各分公司对比",
        "cover": null,
        "url": null,
        "tag": "内部系统"
      },
      {
        "icon": "📞",
        "title": "集团电邀一日五追看板",
        "short": "一日五追",
        "desc": "电邀过程量 · 一日五追",
        "cover": null,
        "url": null,
        "tag": "内部系统"
      },
      {
        "icon": "💻",
        "title": "当日集团业绩大屏",
        "short": "业绩大屏",
        "desc": "集团实时业绩大屏",
        "cover": null,
        "url": null,
        "tag": "内部系统"
      },
      {
        "icon": "⚔️",
        "title": "分公司PK大屏",
        "short": "分公司 PK",
        "desc": "分公司业绩 PK 对比大屏",
        "cover": null,
        "url": null,
        "tag": "内部系统"
      }
    ]
  },
  {
    "name": "分公司",
    "icon": "🏙",
    "layout": "",
    "files": [
      {
        "icon": "📊",
        "title": "分公司数据看板（主看板）",
        "short": "分公司数据看板（主看板）",
        "desc": "活动量总览 · KPI · 部门过程量 · 历史分析",
        "cover": "assets/boards/index.png",
        "url": "../demo/index.html",
        "tag": "在线看板"
      },
      {
        "icon": "🏆",
        "title": "员工三力看板",
        "short": "员工三力看板",
        "desc": "全员三力（业绩 / 过程 / 能力）排名与达标",
        "cover": "assets/boards/sanli.png",
        "url": "../demo/sanli.html",
        "tag": "在线看板"
      },
      {
        "icon": "🧭",
        "title": "网销数据分析",
        "short": "网销数据分析",
        "desc": "网销事业部漏斗 · 过程量 · 人员明细",
        "cover": "assets/boards/ry.png",
        "url": "../demo/ry.html",
        "tag": "在线看板"
      },
      {
        "icon": "📞",
        "title": "电邀数据分析",
        "short": "电邀数据分析",
        "desc": "电邀一日五追 · 邀约转化 · 人员明细",
        "cover": "assets/boards/ry_dy.png",
        "url": "../demo/ry_dy.html",
        "tag": "在线看板"
      },
      {
        "icon": "📒",
        "title": "会计数据分析",
        "short": "会计数据分析",
        "desc": "会计事业部活动量达标 · 周合格率 · 人员明细",
        "cover": "assets/boards/ry_kuaiji.png",
        "url": "../demo/ry_kuaiji.html",
        "tag": "在线看板"
      },
      {
        "icon": "🧭",
        "title": "经营健康度看板",
        "short": "经营健康度",
        "desc": "四大模式统一打分，复合健康分与多级下钻",
        "cover": "assets/boards/health.png",
        "url": "../demo/health.html",
        "tag": "在线看板"
      },
      {
        "icon": "🎧",
        "title": "销售通话录音质检看板",
        "short": "销售通话录音质检看板",
        "desc": "录音 KPI 质检 · 五维评分",
        "cover": "assets/boards/callqa.png",
        "url": "../demo/callqa.html",
        "tag": "在线看板"
      },
      {
        "icon": "🏅",
        "title": "规划师称号在线看板",
        "short": "规划师称号",
        "desc": "称号评定标准与分布",
        "cover": null,
        "url": null,
        "tag": "内部系统"
      }
    ]
  },
  {
    "name": "企业微信在线表格",
    "icon": "📋",
    "layout": "",
    "files": [
      {
        "icon": "👥",
        "title": "见客申诉",
        "short": "见客申诉",
        "desc": "见客量异常申诉 · 在线填写",
        "cover": null,
        "url": null,
        "tag": "内部系统"
      },
      {
        "icon": "📈",
        "title": "活动量报备表",
        "short": "活动量报备",
        "desc": "活动量异常报备 · 在线填写",
        "cover": null,
        "url": null,
        "tag": "内部系统"
      },
      {
        "icon": "☎️",
        "title": "首电申诉表",
        "short": "首电申诉",
        "desc": "首电归属申诉 · 在线填写",
        "cover": null,
        "url": null,
        "tag": "内部系统"
      },
      {
        "icon": "🗂",
        "title": "电审登记表",
        "short": "电审登记",
        "desc": "电审信息登记 · 在线填写",
        "cover": null,
        "url": null,
        "tag": "内部系统"
      }
    ]
  },
  {
    "name": "企业微信在线文档",
    "icon": "📘",
    "layout": "",
    "files": [
      {
        "icon": "📖",
        "title": "呱呱小网管使用手册",
        "short": "网管手册",
        "desc": "系统操作手册 · 使用说明",
        "cover": null,
        "url": null,
        "tag": "内部系统"
      },
      {
        "icon": "🔇",
        "title": "质检异常录音标记无效说明",
        "short": "质检标记说明",
        "desc": "录音质检标记规则 · 无效情形说明",
        "cover": null,
        "url": null,
        "tag": "内部系统"
      }
    ]
  },
  {
    "name": "工具栏",
    "icon": "🧰",
    "layout": "tool",
    "files": [
      {
        "icon": "🔍",
        "title": "百度",
        "short": "百度",
        "desc": "搜索 / 百科 / 翻译",
        "cover": null,
        "url": "https://www.baidu.com",
        "tag": "外部工具"
      },
      {
        "icon": "🧠",
        "title": "DeepSeek",
        "short": "DeepSeek",
        "desc": "深度推理 · 写作 · 免费",
        "cover": null,
        "url": null,
        "tag": "内部系统"
      },
      {
        "icon": "💬",
        "title": "豆包",
        "short": "豆包",
        "desc": "字节 AI 助手 · 对话 / 写作",
        "cover": null,
        "url": "https://www.doubao.com",
        "tag": "外部工具"
      },
      {
        "icon": "🌙",
        "title": "Kimi",
        "short": "Kimi",
        "desc": "长文档阅读 · 论文 / 合同速读",
        "cover": null,
        "url": "https://kimi.moonshot.cn",
        "tag": "外部工具"
      },
      {
        "icon": "🐧",
        "title": "腾讯元宝",
        "short": "腾讯元宝",
        "desc": "腾讯 AI · 公众号 / 微信生态",
        "cover": null,
        "url": "https://yuanbao.tencent.com",
        "tag": "外部工具"
      },
      {
        "icon": "🔎",
        "title": "秘塔 AI 搜索",
        "short": "秘塔 AI 搜索",
        "desc": "无广告搜索 · 带出处 · 深度研究",
        "cover": null,
        "url": null,
        "tag": "内部系统"
      },
      {
        "icon": "📄",
        "title": "WPS 云文档",
        "short": "WPS 云文档",
        "desc": "在线 Word / Excel / PPT",
        "cover": null,
        "url": null,
        "tag": "内部系统"
      },
      {
        "icon": "📝",
        "title": "腾讯文档",
        "short": "腾讯文档",
        "desc": "多人协作表格 · 文档",
        "cover": null,
        "url": null,
        "tag": "内部系统"
      },
      {
        "icon": "🏢",
        "title": "企查查",
        "short": "企查查",
        "desc": "查企业工商 / 股东 / 风险",
        "cover": null,
        "url": null,
        "tag": "内部系统"
      },
      {
        "icon": "🔳",
        "title": "草料二维码",
        "short": "草料二维码",
        "desc": "二维码生成 · 表单 · 活码",
        "cover": null,
        "url": null,
        "tag": "内部系统"
      },
      {
        "icon": "🗺",
        "title": "高德地图",
        "short": "高德地图",
        "desc": "导航 / 路线规划",
        "cover": null,
        "url": "https://www.amap.com",
        "tag": "外部工具"
      },
      {
        "icon": "🚄",
        "title": "铁路 12306",
        "short": "铁路 12306",
        "desc": "火车票查询 / 购票",
        "cover": null,
        "url": "https://www.12306.cn",
        "tag": "外部工具"
      }
    ]
  }
],

  notices: [
  {
    "date": "2026-10-27",
    "time": "08:02",
    "cat": "brief",
    "title": "2026 年 10 月 27 日数据简报 TV，星期二，第 1720 期",
    "content": [
      "示例简报正文。演示版已移除原始内容，此处仅用于展示列表与翻页交互。",
      "【演示数据】本栏目用于说明「每日简报」板块的呈现方式。"
    ]
  },
  {
    "date": "2026-10-26",
    "time": "08:05",
    "cat": "brief",
    "title": "2026 年 10 月 26 日数据简报 TV，星期一，第 1719 期",
    "content": [
      "示例简报正文。演示版已移除原始内容。"
    ]
  },
  {
    "date": "2026-10-25",
    "time": "08:10",
    "cat": "brief",
    "title": "2026 年 10 月 25 日数据简报 TV，星期日，第 1718 期",
    "content": [
      "示例简报正文。演示版已移除原始内容。"
    ]
  },
  {
    "date": "2026-10-20",
    "time": "14:24",
    "cat": "notice",
    "title": "关于 10 月考核口径调整的通知（演示）",
    "content": [
      "本通知用于演示「重要通知」弹窗的展示效果。",
      "1、口径调整说明示例文本。",
      "2、生效时间与适用范围示例文本。"
    ]
  },
  {
    "date": "2026-10-19",
    "time": "18:05",
    "cat": "notice",
    "title": "数据录入操作手册已更新（演示）",
    "content": [
      "操作手册更新说明示例文本。"
    ]
  },
  {
    "date": "2026-10-19",
    "time": "14:17",
    "cat": "notice",
    "title": "季度数据核对工作提醒（演示）",
    "content": [
      "工作提醒示例文本。"
    ]
  },
  {
    "date": "2026-10-19",
    "time": "14:04",
    "cat": "notice",
    "title": "11 月激励方案发布（演示）",
    "content": [
      "激励方案说明示例文本。"
    ]
  },
  {
    "date": "2026-10-19",
    "time": "13:57",
    "cat": "notice",
    "title": "关于月度账单核对时间节点的通知（演示）",
    "content": [
      "时间节点说明示例文本。"
    ]
  }
],

  boards: [
    { id: "nav",     name: "看板导航页",     en: "Portal",           group: "入口",   desc: "六大分区统一入口，侧栏定位、搜索、公告与轮播", path: "../demo/nav.html",     thumb: null },
    { id: "index",   name: "活动量看板",     en: "Activity Board",   group: "核心",   desc: "KPI 总览 · 部门过程量 · 全员明细 · 历史趋势",   path: "../demo/index.html",   thumb: null },
    { id: "sanli",   name: "员工三力",       en: "Three-Force",      group: "组织",   desc: "英雄 / 标准 / 破零 / 无效四层人力分级",         path: "../demo/sanli.html",   thumb: null },
    { id: "health",  name: "经营健康度",     en: "Org Health",       group: "组织",   desc: "四大模式统一打分，复合健康分与多级下钻",          path: "../demo/health.html",  thumb: null },
    { id: "callqa",  name: "通话质检评分",   en: "Call QA",          group: "质检",   desc: "827 条录音五维评分与改进建议",                    path: "../demo/callqa.html",  thumb: null },
    { id: "ry",      name: "网销数据分析",   en: "Online Sales",     group: "过程量", desc: "网销漏斗逐人逐环节拆解",                          path: "../demo/ry.html",      thumb: null },
    { id: "ry_dy",   name: "电邀数据分析",   en: "Outbound",         group: "过程量", desc: "通时 / 拨打 / 接通三率联动",                      path: "../demo/ry_dy.html",   thumb: null },
    { id: "ry_kuaiji", name: "会计数据分析", en: "Accounting",       group: "过程量", desc: "财务侧活动量考核与票据抽检",                      path: "../demo/ry_kuaiji.html", thumb: null },
    { id: "yj",      name: "业绩含续费",     en: "Revenue",          group: "结果",   desc: "新签与续费合并口径，三级下钻",                    path: "../demo/yj.html",      thumb: null },
    { id: "sku",     name: "订单商品明细",   en: "Order SKU",        group: "结果",   desc: "订单级明细穿透查询",                              path: "../demo/sku.html",     thumb: null }
  ]
};
