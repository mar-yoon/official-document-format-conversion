const { apiBaseUrl } = require('../../config')

Page({
  data: { filePath: '', fileName: '', fileSizeLabel: '', busy: false, error: '', report: null, downloadPath: '' },
  chooseFile() {
    wx.chooseMessageFile({ count: 1, type: 'file', extension: ['docx'], success: res => {
      const file = res.tempFiles[0]
      if (!file || !/\.docx$/i.test(file.name)) return this.setData({ error: '请选择 .docx Word 文档' })
      if (file.size > 9 * 1024 * 1024) return this.setData({ error: '文档不能超过 9 MB' })
      this.setData({ filePath: file.path, fileName: file.name, fileSizeLabel: `${(file.size / 1024).toFixed(1)} KB`, error: '', report: null, downloadPath: '' })
    } })
  },
  formatDocument() {
    if (!this.data.filePath || this.data.busy) return
    if (apiBaseUrl.includes('your-domain.example.com')) return this.setData({ error: '请先在 config.js 中配置服务器 HTTPS 域名' })
    this.setData({ busy: true, error: '', report: null })
    wx.uploadFile({
      url: `${apiBaseUrl}/api/jobs?template=gov_official`, filePath: this.data.filePath, name: 'file',
      success: res => {
        let body
        try { body = JSON.parse(res.data) } catch (_) { body = {} }
        if (res.statusCode !== 200 || !body.download_path) {
          this.setData({ error: body.error || '处理失败，请稍后重试', busy: false }); return
        }
        this.setData({ report: body.report, downloadPath: body.download_path, busy: false })
        wx.showToast({ title: '排版完成', icon: 'success' })
      },
      fail: () => this.setData({ error: '上传失败，请检查网络和服务器域名设置', busy: false })
    })
  },
  openResult() {
    if (!this.data.downloadPath) return
    wx.showLoading({ title: '正在打开文档' })
    wx.downloadFile({
      url: `${apiBaseUrl}${this.data.downloadPath}`,
      success: res => {
        if (res.statusCode !== 200) { wx.hideLoading(); this.setData({ error: '下载链接已过期，请重新处理文档' }); return }
        wx.openDocument({ filePath: res.tempFilePath, fileType: 'docx', showMenu: true,
          success: () => wx.hideLoading(),
          fail: () => { wx.hideLoading(); this.setData({ error: '无法打开文档，请重试' }) }
        })
      },
      fail: () => { wx.hideLoading(); this.setData({ error: '下载失败，请检查网络' }) }
    })
  }
})
