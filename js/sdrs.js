/* SDR 接收机配置文件 — 编辑本文件即可更改站点的默认接收机列表。
   字段:
     name  显示名称
     url   接收机地址（http/https 完整 URL）
     type  kiwi | openwebrx | websdr（也可省略，自动按地址识别）
     def   true = 默认接收机（卡片 🎧 按钮直接打开的那台），仅设一台
   更新说明:
     - 修改 sdrs 数组后，请把 rev 加一（如 1 -> 2），
       并同步把 index.html 里 <script src="js/sdrs.js?v=2"> 的 ?v= 改成相同数字（防浏览器缓存）。
       rev 变化后，所有访客下次打开页面会自动采用新列表；
       若访客曾在页面上自行增删过接收机，其自定义仍然优先生效，
       直到在"🎧 SDR 收听"弹窗中点击"↺ 恢复默认"重新同步本文件。
*/
window.RADIO_SDRS_CFG = {
  rev: 2,
  sdrs: [
    { name: 'KiwiSDR JK1LOT', url: 'http://jk1lot.ham-radio-op.net:8073/', type: 'kiwi', def: true },
    { name: 'KiwiSDR jj8ntm', url: 'http://jj8ntm.proxy.kiwisdr.com:8073/', type: 'kiwi', def: false },
    { name: 'KiwiSDR jjm0311', url: 'http://jjm0311.proxy.kiwisdr.com:8073/', type: 'kiwi', def: false },
    // 示例（去掉行首注释即可启用；type 可省略自动识别）:
    // { name: 'OpenWebRX Twente', url: 'https://rx.example.org/', type: 'openwebrx' },
    // { name: 'WebSDR Twente', url: 'http://websdr.ewi.utwente.nl:8901/', type: 'websdr' },
  ],
};
