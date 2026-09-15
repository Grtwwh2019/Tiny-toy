package mailtool;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.util.*;
import java.util.concurrent.TimeUnit;

final class Store {
    final Path dir;
    Store(Path dir)throws IOException {this.dir=dir;Files.createDirectories(dir);}
    static Path defaultDir(){
        String appData=System.getenv("APPDATA");
        return appData!=null ? Paths.get(appData,"OutlookTemplateTool") : Paths.get(System.getProperty("user.home"),".outlook-template-tool");
    }
    static void atomic(Path path,byte[] data)throws IOException {
        Path tmp=Files.createTempFile(path.getParent(),"save-",".tmp");
        try {
            Files.write(tmp,data);
            try{Files.move(tmp,path,StandardCopyOption.ATOMIC_MOVE,StandardCopyOption.REPLACE_EXISTING);}
            catch(AtomicMoveNotSupportedException ex){Files.move(tmp,path,StandardCopyOption.REPLACE_EXISTING);}
        }finally{Files.deleteIfExists(tmp);}
    }
    List<Core.Template> loadTemplates()throws IOException {
        Path file=dir.resolve("templates.json");
        if(!Files.exists(file)){
            List<Core.Template> sample=new ArrayList<>();
            sample.add(new Core.Template(UUID.randomUUID().toString(),"付款提醒","{{收件邮箱}}","{{姓名}} · 付款提醒 · {{今天:yyyy-MM-dd}}","{{姓名}}，您好：\n\n本次待付金额为 {{金额}} 元，请于 {{今天+7天:yyyy-MM-dd}} 前安排付款。\n\n如已付款，请忽略此提醒。谢谢！\n\n{{署名}}"));
            saveTemplates(sample);return sample;
        }
        Object parsed=Json.parse(new String(Files.readAllBytes(file),StandardCharsets.UTF_8));
        if(!(parsed instanceof List))throw new IOException("模板文件格式错误");
        List<Core.Template> result=new ArrayList<>();Set<String> ids=new HashSet<>();
        for(Object o:(List<?>)parsed){Core.Template t=Core.Template.from(o);if(!ids.add(t.id))throw new IOException("模板编号重复");result.add(t);}
        return result;
    }
    void saveTemplates(List<Core.Template> templates)throws IOException {
        List<Object> data=new ArrayList<>();for(Core.Template t:templates)data.add(t.json());
        atomic(dir.resolve("templates.json"),Json.write(data).getBytes(StandardCharsets.UTF_8));
    }
    Properties loadSettings()throws IOException {
        Properties p=new Properties();Path file=dir.resolve("settings.properties");
        if(Files.exists(file))try(InputStream in=Files.newInputStream(file)){p.load(in);}return p;
    }
    void saveSettings(Properties p)throws IOException {
        ByteArrayOutputStream b=new ByteArrayOutputStream();p.store(b,"Outlook Template Tool");atomic(dir.resolve("settings.properties"),b.toByteArray());
    }
    static boolean windows(){return System.getProperty("os.name").toLowerCase(Locale.ROOT).contains("windows");}
    String readRefresh(String key)throws Exception {
        Path p=dir.resolve("session.dpapi");if(!windows()||!Files.exists(p))return "";
        Map<String,Object> m=Json.object(Json.parse(new String(protect(Files.readAllBytes(p),false),StandardCharsets.UTF_8)));
        return key.equals(Json.str(m,"key"))?Json.str(m,"refresh"):"";
    }
    void saveRefresh(String key,String refresh)throws Exception {
        if(!windows())return;
        byte[] plain=Json.write(Json.obj("key",key,"refresh",refresh)).getBytes(StandardCharsets.UTF_8);
        try{atomic(dir.resolve("session.dpapi"),protect(plain,true));}finally{Arrays.fill(plain,(byte)0);}
    }
    void clearRefresh()throws IOException {Files.deleteIfExists(dir.resolve("session.dpapi"));}
    // DPAPI ties the cache to the current Windows user. Token bytes travel via stdin,
    // never command-line arguments, shell interpolation, logs or plaintext files.
    private byte[] protect(byte[] bytes,boolean encrypt)throws Exception {
        String root=System.getenv("SystemRoot");if(root==null)throw new IOException("找不到 Windows 系统目录");
        String ps=Paths.get(root,"System32","WindowsPowerShell","v1.0","powershell.exe").toString();
        String script="$ErrorActionPreference='Stop'; Add-Type -AssemblyName System.Security; "
            +"$b=[Convert]::FromBase64String([Console]::In.ReadToEnd()); "
            +"$r=[Security.Cryptography.ProtectedData]::"+(encrypt?"Protect":"Unprotect")
            +"($b,$null,[Security.Cryptography.DataProtectionScope]::CurrentUser); "
            +"[Console]::Out.Write([Convert]::ToBase64String($r))";
        Process process=new ProcessBuilder(ps,"-NoLogo","-NoProfile","-NonInteractive","-Command",script).redirectError(ProcessBuilder.Redirect.INHERIT).start();
        try{
            try(OutputStream out=process.getOutputStream()){out.write(Base64.getEncoder().encode(bytes));}
            final ByteArrayOutputStream result=new ByteArrayOutputStream();
            Thread reader=new Thread(()->{try{copy(process.getInputStream(),result);}catch(IOException ignored){}},"credential-cache");reader.setDaemon(true);reader.start();
            if(!process.waitFor(15,TimeUnit.SECONDS)){process.destroyForcibly();throw new IOException("Windows 安全存储超时");}
            reader.join(1000);if(process.exitValue()!=0||reader.isAlive())throw new IOException("Windows 安全存储不可用");
            return Base64.getDecoder().decode(result.toString("US-ASCII").trim());
        }finally{process.destroy();}
    }
    static void copy(InputStream in,OutputStream out)throws IOException {
        byte[] b=new byte[8192];int n;while((n=in.read(b))!=-1)out.write(b,0,n);
    }
}
