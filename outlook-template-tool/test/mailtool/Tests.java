package mailtool;

import java.io.*;
import java.net.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.time.LocalDate;
import java.util.*;
import java.util.concurrent.atomic.*;

public final class Tests {
    static int checks;
    static void check(boolean ok,String label){checks++;if(!ok)throw new AssertionError(label);}
    interface Checked {void run()throws Exception;}
    static void rejects(Checked action,String label)throws Exception {
        try{action.run();throw new AssertionError("Expected rejection: "+label);}catch(IllegalArgumentException|IOException ex){checks++;}
    }
    public static void main(String[] args)throws Exception {
        LocalDate leap=LocalDate.of(2024,2,28);
        check(Core.date("今天+1天:yyyy-MM-dd",leap).equals("2024-02-29"),"leap day");
        check(Core.date("今天+2天:yyyy-MM-dd",leap).equals("2024-03-01"),"month boundary");
        check(Core.date("今天-1天:yyyy年M月d日",LocalDate.of(2026,1,1)).equals("2025年12月31日"),"year boundary");
        rejects(()->Core.date("今天+999999999999999999天",leap),"overflow");
        rejects(()->Core.date("今天:bad[",leap),"invalid date pattern");
        rejects(()->Core.tokens("{{姓名}"),"unclosed placeholder");
        rejects(()->Core.tokens("{{}}"),"empty placeholder");

        Core.Template t=new Core.Template("one","Test","{{姓名}}<{{邮箱}}>","财务部 <finance@example.com>","{{姓名}} / {{今天}}","{{姓名}} owes {{金额}}. {{今天+7天:yyyy-MM-dd}}");
        check(Core.variables(t).equals(new LinkedHashSet<>(Arrays.asList("姓名","邮箱","金额"))),"variables include address fields");
        Map<String,String> v=new HashMap<>();v.put("邮箱","me@example.com");v.put("姓名","王$\\同学{{今天}}");v.put("金额","1,250.50");
        Core.Message msg=Core.generate(t,v,leap);
        check(msg.body.startsWith("王$\\同学{{今天}} owes 1,250.50"),"literal replacement, no recursive execution");
        check(msg.body.endsWith("2024-03-06"),"date snapshot");
        check(msg.recipients.get(0).display().equals("王$\\同学{{今天}} <me@example.com>"),"named recipient rendered");
        check(msg.cc.get(0).display().equals("财务部 <finance@example.com>"),"named cc rendered");
        check(msg.preview("sender@example.com").contains("抄送人：财务部 <finance@example.com>"),"cc preview");
        Map<String,String> missing=new HashMap<>(v);missing.remove("金额");rejects(()->Core.generate(t,missing,leap),"missing variable");

        Core.Message named=new Core.Message("Jim zheng<jim_zheng@mail.com>; Jane <jane@example.com>","Boss <boss@example.com>；audit@example.com","subject","body");
        check(named.recipients.size()==2&&named.cc.size()==2,"English and Chinese semicolon lists");
        check(named.recipients.get(0).name.equals("Jim zheng")&&named.recipients.get(0).email.equals("jim_zheng@mail.com"),"compact display address");
        Core.Message deduped=new Core.Message("Jim <SAME@example.com>; same@example.com","","subject","body");
        check(deduped.recipients.size()==1,"case-insensitive duplicate removed");
        rejects(()->new Core.Message("a@b.com\r\nBcc:x@y.com","","subject","body"),"recipient injection");
        rejects(()->new Core.Message("a@b.com","x@y.com\nBcc:z@y.com","subject","body"),"cc injection");
        rejects(()->new Core.Message("a@b.com","","subject\nBcc:x@y.com","body"),"subject injection");
        rejects(()->new Core.Message("Jim <broken>","","subject","body"),"invalid named mailbox");
        rejects(()->new Core.Message("<a@b.com>","","subject","body"),"empty display name");
        rejects(()->new Core.Message("a@b.com;","","subject","body"),"empty address");

        Core.Message unicode=new Core.Message("one@example.com；two@example.com","抄送 <copy@example.com>","金额 & # + ?","你好\n第二行 & 100%");
        Map<String,String> params=query(unicode.mailto(true).getRawSchemeSpecificPart().split("\\?",2)[1]);
        check(params.get("subject").equals(unicode.subject),"subject URI round trip");
        check(params.get("body").equals("你好\r\n第二行 & 100%"),"body URI UTF-8 and CRLF");
        check(params.get("cc").equals("copy@example.com"),"mailto cc");
        check(unicode.mailto(true).toString().startsWith("mailto:one%40example.com,two%40example.com"),"mailto uses SMTP addresses");
        String longBody=String.join("",Collections.nCopies(3000,"长"));Core.Message longMail=new Core.Message("one@example.com","","test",longBody);
        rejects(()->longMail.mailto(true),"no long-body truncation");check(longMail.mailto(false).toString().length()<Core.MAILTO_LIMIT,"manual paste fallback");

        Map<String,Object> payload=Json.object(Json.parse(Json.write(named.outlookPayload("open"))));
        check(Json.str(payload,"operation").equals("open"),"Outlook operation");
        check(((List<?>)payload.get("to")).get(0).equals("Jim zheng <jim_zheng@mail.com>"),"Outlook receives display name");
        check(((List<?>)payload.get("cc")).size()==2,"Outlook receives cc list");
        String odd="quote\" slash\\ control\n\t\b\u0000 中文 😀";check(odd.equals(Json.parse(Json.write(odd))),"JSON escapes round trip");
        rejects(()->Json.parse("{}junk"),"JSON trailing garbage");rejects(()->Json.parse("[1,]"),"JSON trailing comma");

        Path path=Files.createTempDirectory("mailtool-test-");Store store=new Store(path);
        check(store.loadTemplates().size()==1,"first-run sample");check(store.loadTemplates().get(0).cc.isEmpty(),"sample cc optional");
        store.saveTemplates(Arrays.asList(t));check(store.loadTemplates().get(0).cc.equals(t.cc),"template cc persistence");
        Map<String,Object> legacy=Json.obj("id","old","name","Old","to","a@b.com","subject","s","body","b");
        Store.atomic(path.resolve("templates.json"),Json.write(Arrays.asList(legacy)).getBytes(StandardCharsets.UTF_8));
        check(store.loadTemplates().get(0).cc.isEmpty(),"old template migrates with blank cc");
        Properties settings=new Properties();settings.setProperty("direct","true");store.saveSettings(settings);check(store.loadSettings().getProperty("direct").equals("true"),"settings persistence");
        Files.write(path.resolve("session.dpapi"),new byte[]{1,2,3});store.clearLegacyCredentials();check(!Files.exists(path.resolve("session.dpapi")),"legacy login token removed");
        store.saveTemplates(Collections.emptyList());check(store.loadTemplates().isEmpty(),"empty list stays empty");
        Files.write(path.resolve("templates.json"),"broken".getBytes(StandardCharsets.UTF_8));rejects(()->store.loadTemplates(),"corrupt data not overwritten");
        check(new String(Files.readAllBytes(path.resolve("templates.json")),StandardCharsets.UTF_8).equals("broken"),"preserve corrupt data");

        testOutlook(named);
        System.out.println("PASS: "+checks+" checks. Outlook automation was mocked; no real email was opened or sent.");
    }

    static void testOutlook(Core.Message message)throws Exception {
        {
            AtomicInteger runs=new AtomicInteger(),sends=new AtomicInteger();AtomicReference<Map<String,Object>> last=new AtomicReference<>();
            Outlook outlook=new Outlook((input,sending)->{
                runs.incrementAndGet();Map<String,Object> data=Json.object(Json.parse(input));last.set(data);if(sending)sends.incrementAndGet();
                return new Outlook.Result(0,"{\"account\":\"sender@example.com\"}","");
            });
            check(outlook.account().equals("sender@example.com"),"classic Outlook current account");
            check(outlook.open(message).equals("sender@example.com"),"classic Outlook open");check(Json.str(last.get(),"operation").equals("open"),"open operation dispatched");
            check(((List<?>)last.get().get("cc")).size()==2,"cc dispatched to Outlook");
            check(outlook.send(message).equals("sender@example.com"),"classic Outlook send");check(sends.get()==1&&runs.get()==3,"send invoked exactly once");

            Outlook unavailable=new Outlook((input,sending)->new Outlook.Result(4,"","OUTLOOK_UNAVAILABLE: Class not registered"));
            try{unavailable.account();throw new AssertionError("Expected classic Outlook unavailable");}catch(Outlook.NotAvailableException ex){checks++;}
            AtomicInteger attempts=new AtomicInteger();Outlook uncertain=new Outlook((input,sending)->{attempts.incrementAndGet();throw new IOException("timeout");});
            rejects(()->uncertain.send(message),"uncertain send surfaced");check(attempts.get()==1,"uncertain send not retried");
            Outlook rejected=new Outlook((input,sending)->new Outlook.Result(3,"","OUTLOOK_ERROR: access denied"));
            rejects(()->rejected.send(message),"Outlook send rejection surfaced");
        }
    }

    static Map<String,String> query(String value)throws Exception {
        Map<String,String> result=new HashMap<>();for(String part:value.split("&")){String[] pair=part.split("=",2);result.put(URLDecoder.decode(pair[0],"UTF-8"),pair.length==2?URLDecoder.decode(pair[1],"UTF-8"):"");}return result;
    }
}
