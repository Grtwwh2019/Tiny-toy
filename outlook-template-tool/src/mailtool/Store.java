package mailtool;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.util.*;

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
            sample.add(new Core.Template(UUID.randomUUID().toString(),"付款提醒","{{收件人姓名}}<{{收件邮箱}}>","","{{收件人姓名}} · 付款提醒 · {{今天:yyyy-MM-dd}}","{{收件人姓名}}，您好：\n\n本次待付金额为 {{金额}} 元，请于 {{今天+7天:yyyy-MM-dd}} 前安排付款。\n\n如已付款，请忽略此提醒。谢谢！\n\n{{署名}}"));
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
    void clearLegacyCredentials()throws IOException {Files.deleteIfExists(dir.resolve("session.dpapi"));}
    static void copy(InputStream in,OutputStream out)throws IOException {
        byte[] b=new byte[8192];int n;while((n=in.read(b))!=-1)out.write(b,0,n);
    }
}
