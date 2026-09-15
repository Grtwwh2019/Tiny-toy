package mailtool;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.util.*;
import java.util.concurrent.TimeUnit;

/** Controls the locally configured classic Outlook profile through its COM object model. */
final class Outlook {
    interface Runner {Result run(String input,boolean sending)throws Exception;}
    static final class Result {
        final int status;final String output,error;
        Result(int status,String output,String error){this.status=status;this.output=output;this.error=error;}
    }
    static final class NotAvailableException extends IOException {
        private static final long serialVersionUID=1L;
        NotAvailableException(String message){super(message);}
    }

    private static final String SCRIPT=
        "$ErrorActionPreference='Stop'; "
        +"[Console]::InputEncoding=New-Object Text.UTF8Encoding($false); "
        +"[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false); "
        +"try {$outlook=New-Object -ComObject Outlook.Application} catch {[Console]::Error.Write('OUTLOOK_UNAVAILABLE: '+$_.Exception.Message);exit 4}; "
        +"try { "
        +"$data=ConvertFrom-Json ([Console]::In.ReadToEnd()); "
        +"$session=$outlook.Session; "
        +"if($session.Accounts.Count -lt 1){throw '当前 Outlook 没有可用账号'}; "
        +"$account=$session.Accounts.Item(1); "
        +"$smtp=[string]$account.SmtpAddress; if([string]::IsNullOrWhiteSpace($smtp)){$smtp=[string]$session.CurrentUser.Name}; "
        +"if($data.operation -eq 'status'){[Console]::Out.Write((@{account=$smtp;version=[string]$outlook.Version}|ConvertTo-Json -Compress));exit 0}; "
        +"$mail=$outlook.CreateItem(0); "
        +"$mail.SendUsingAccount=$account; "
        +"$mail.To=(@($data.to)-join '; '); $mail.CC=(@($data.cc)-join '; '); "
        +"$mail.Subject=[string]$data.subject; $mail.Body=[string]$data.body; "
        +"if(-not $mail.Recipients.ResolveAll()){throw 'Outlook 无法解析一个或多个收件人/抄送人'}; "
        +"if($data.operation -eq 'open'){[void]$mail.Display($false)} "
        +"elseif($data.operation -eq 'send'){$mail.Send()} else {throw '未知操作'}; "
        +"[Console]::Out.Write((@{account=$smtp}|ConvertTo-Json -Compress)); "
        +"} catch { [Console]::Error.Write('OUTLOOK_ERROR: '+$_.Exception.Message); exit 3 }";

    private final Runner runner;private final boolean checkPlatform;
    Outlook(){this.runner=new PowerShellRunner();this.checkPlatform=true;}
    Outlook(Runner runner){this.runner=runner;this.checkPlatform=false;}

    String account()throws Exception{return invoke(Json.obj("operation","status"),false,"检测经典 Outlook");}
    String open(Core.Message message)throws Exception{return invoke(message.outlookPayload("open"),false,"打开 Outlook 邮件");}
    String send(Core.Message message)throws Exception{return invoke(message.outlookPayload("send"),true,"发送 Outlook 邮件");}

    private String invoke(Map<String,Object> input,boolean sending,String action)throws Exception {
        if(checkPlatform&&!Store.windows())throw new NotAvailableException("经典 Outlook 本机接口仅在 Windows 上可用");
        final Result r;
        try{r=runner.run(Json.write(input),sending);}catch(NotAvailableException ex){throw ex;}
        catch(Exception ex){
            if(sending)throw new IOException("发送结果尚不确定。请先检查 Outlook 的“已发送邮件”，再决定是否重试，避免重复发送。",ex);
            throw new IOException(action+"失败："+safe(ex.getMessage()),ex);
        }
        if(r.status!=0){
            String detail=safe(r.error).replace("OUTLOOK_ERROR:","").trim();
            if(detail.isEmpty())detail="未检测到可用的经典 Outlook 配置";
            if(r.status==4)throw new NotAvailableException("未检测到经典 Outlook 本机接口。新版 Outlook 不支持直接发送。详情："+detail.replace("OUTLOOK_UNAVAILABLE:","").trim());
            if(sending)throw new IOException("Outlook 未完成发送："+detail+"。请检查“已发送邮件”后再决定是否重试。");
            throw new IOException(action+"失败："+detail);
        }
        Map<String,Object> response=Json.object(Json.parse(r.output.trim()));String account=Json.str(response,"account");
        if(account.trim().isEmpty())throw new IOException("Outlook 未返回当前发件账号");return account;
    }

    private static String safe(String value){return value==null?"":value.replace('\r',' ').replace('\n',' ').trim();}

    private static final class PowerShellRunner implements Runner {
        public Result run(String input,boolean sending)throws Exception {
            String root=System.getenv("SystemRoot");if(root==null)throw new NotAvailableException("找不到 Windows 系统目录");
            String executable=Paths.get(root,"System32","WindowsPowerShell","v1.0","powershell.exe").toString();
            String encoded=Base64.getEncoder().encodeToString(SCRIPT.getBytes("UTF-16LE"));
            Process process=new ProcessBuilder(executable,"-NoLogo","-NoProfile","-NonInteractive","-EncodedCommand",encoded).start();
            ByteArrayOutputStream stdout=new ByteArrayOutputStream(),stderr=new ByteArrayOutputStream();
            Thread out=reader(process.getInputStream(),stdout,"outlook-output"),err=reader(process.getErrorStream(),stderr,"outlook-error");
            try{
                try(OutputStream stdin=process.getOutputStream()){stdin.write(input.getBytes(StandardCharsets.UTF_8));}
                if(!process.waitFor(90,TimeUnit.SECONDS)){
                    process.destroyForcibly();
                    throw new IOException(sending?"等待 Outlook 发送超时":"等待 Outlook 响应超时");
                }
                out.join(2000);err.join(2000);
                if(out.isAlive()||err.isAlive())throw new IOException("读取 Outlook 响应超时");
                return new Result(process.exitValue(),stdout.toString("UTF-8"),stderr.toString("UTF-8"));
            }finally{process.destroy();}
        }
        private static Thread reader(InputStream input,OutputStream output,String name){
            Thread t=new Thread(()->{try{Store.copy(input,output);}catch(IOException ignored){}},name);t.setDaemon(true);t.start();return t;
        }
    }
}
