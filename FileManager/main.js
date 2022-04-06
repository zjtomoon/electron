const {app,BrowserWindow} = require('electron')

const isDev = require('electron-is-dev')

let mainWindow

app.on('ready',() => {
    mainWindow = new BrowserWindow({
        width:1024,
        height:680,
        webPreferences:{
            nodeIntegration:true,
        }
    })

    const urlLocation = isDev ? 'http://localhost:3000' : 'dummyurl'
    mainWindow.loadURL(urlLocation)
    mainWindow.webContents.openDevTools()
})

// 参考：https://www.jianshu.com/p/74af504a1d3d
