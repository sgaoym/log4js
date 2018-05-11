/**
 * @version 0.1
 * @author sgaoym sgaoym@gmail.com
 */
if(!log){
    var log = {};
    log.fso     = new ActiveXObject("Scripting.FileSystemObject");
    log.path    = "d://weblog/";    //日志路径 
    log.runFile = "c://run";        //运行日志开启文件
    log.dbgFile = "c://dbg";        //调试日志开启文件
    log.errFile = "c://err";        //错误日志开启文件
    log.mode    = 8;                //日志模式
    /**
     * 设置日志路径
     */
    log.setPath = function(path){
        if(path === undefined || path === null ||　log.fso === null){return false;}
        try{
            if(!log.fso.FolderExists(path)){
                log.createFolder(path);
            }
        }catch(e){
            return false;
        }
        log.path = path;
        return true;
    };
    /**
     * 获得日志级别
     */
    log.getLevel = function(){
        if(log.fso === null) {return false;}
        var le = 0;
        if (log.fso.FileExists(log.runFile)){le = 1;}
        if (log.fso.FileExists(log.dbgFile)){le = 3;}
        if (log.fso.FileExists(log.errFile)){le = 11;}
        return le;
    };
    /**
     * 输出运行日志
     */
    log.run = function(data){
        if(log.level < 1) {return false;}
        try{
            var tf = log.fso.OpenTextFile(log.path+log.file, log.mode, true);
            data = log.formatDateTime(new Date()) + ' ' + data;
            tf.WriteLine(data);
            tf.Close();
        }catch(e){
            return false;
        }
        return true;
    };
    /**
     * 输出调试日志
     */
    log.dbg = function(data){
        if(log.level < 3) {return false;}
        try{
            var tf = log.fso.OpenTextFile(log.path+log.file, log.mode, true);
            tf.Write(data);
            tf.Close();
        }catch(e){
            return false;
        }
        return true;
    };
    /**
     * 输出错误日志
     */
    log.err = function(data){
        if(log.level < 11) {return false;}
        try{
            var tf = log.fso.OpenTextFile(log.path+log.file, log.mode, true);
            tf.Write(data);
            tf.Close();
        }catch(e){
            return false;
        }
        return true;
    };
    /**
     * 格式化时间
     */
    log.formatDateTime =  function(dt){
        var y = dt.getFullYear();
        var m = dt.getMonth() + 1;
        var d = dt.getDate();
        var h = dt.getHours(); 
        var i = dt.getMinutes(); 
        var s = dt.getSeconds(); 
        m = m < 10 ? '0' + m : m;
        d = d < 10 ? '0' + d : d;
        h = h < 10 ? '0' + h : h;
        i = i < 10 ? '0' + i : i;
        s = s < 10 ? '0' + s : s;
        return y + '-' + m + '-' + d + " "+ h +":"+ i +":"+ s;
    };
    /**
     * 创建多级文件夹
     */
    log.createFolder = function(fd){
        var index = -1, i = 0;
        while((index = fd.indexOf('/', i)) != -1){
            if(!log.fso.FolderExists(fd.substring(0, index))){
                log.fso.CreateFolder(fd.substring(0, index));
            }
            i = index + 1;   
        }
    };
    log.setPath(log.path);
    log.level= log.getLevel();
    log.file = 'log_'+log.formatDateTime(new Date()).substring(0, 10)+'.log'; 
}