const { app, BrowserWindow } = require('electron');

const createWindow = () =>{
    const window = new BrowserWindow({
        with: 800,
        height: 500,
        minWidth: 400,
        minHeight: 400,
        webPreferences: {
            nodeIntegration: true,
            contextIsolation: false,
        },
    })

    window.loadFile('index.html')
}

app.whenReady().then(() => {
    createWindow()
})