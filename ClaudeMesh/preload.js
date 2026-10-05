const { contextBridge, ipcRenderer, webUtils } = require('electron');

contextBridge.exposeInMainWorld('api', {
  onTelemetry: cb => ipcRenderer.on('telemetry', (_e, s) => cb(s)),
  pty: {
    spawn: opts => ipcRenderer.invoke('pty:spawn', opts),
    write: (id, d) => ipcRenderer.send('pty:write', id, d),
    resize: (id, c, r) => ipcRenderer.send('pty:resize', id, c, r),
    kill: id => ipcRenderer.send('pty:kill', id),
    onData: cb => ipcRenderer.on('pty:data', (_e, id, d) => cb(id, d)),
    onExit: cb => ipcRenderer.on('pty:exit', (_e, id, code) => cb(id, code)),
  },
  ext: {
    send: (tty, text) => ipcRenderer.invoke('ext:send', tty, text),
    interrupt: tty => ipcRenderer.invoke('ext:interrupt', tty),
    focus: tty => ipcRenderer.invoke('ext:focus', tty),
  },
  ctl: { send: (id, text, submit) => ipcRenderer.invoke('ctl:send', id, text, submit) },
  pickDir: () => ipcRenderer.invoke('dlg:dir'),
  notify: (t, b) => ipcRenderer.invoke('notify', t, b),
  home: () => ipcRenderer.invoke('home'),
  clipRead: () => ipcRenderer.invoke('clip:read'),
  listCommands: cwd => ipcRenderer.invoke('cmds:list', cwd),
  pathForFile: f => webUtils.getPathForFile(f),
  history: () => ipcRenderer.invoke('hist:list'),
  rename: (sid, name) => ipcRenderer.invoke('hist:rename', sid, name),
  agents: {
    list: () => ipcRenderer.invoke('agents:list'),
    convert: (sid, cwd, task) => ipcRenderer.invoke('agent:convert', sid, cwd, task),
    stop: id => ipcRenderer.invoke('agent:stop', id),
    logs: id => ipcRenderer.invoke('agent:logs', id),
  },
  refreshUsage: () => ipcRenderer.invoke('usage:refresh'),
  debugShot: f => ipcRenderer.invoke('debug:shot', f),
});
