package mailtool;

import java.io.*;
import java.net.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.time.LocalDate;
import java.util.*;
import java.util.concurrent.atomic.*;
import com.sun.net.httpserver.HttpServer;

public final class Tests {
    static int checks;
    static void check(boolean ok,String label){checks++;if(!ok)throw new AssertionError(label);}
    interface Checked {void run()throws Exception;}
    static void rejects(Checked action,String label)throws Exception {
        try{action.run();throw new AssertionError("Expected rejection: "+label);}catch(IllegalArgumentException|IOException ex){checks++;}
    }
    public static void main(String[] args)throws Exception {
        System.setProperty("sun.net.http.retryPost","false");
        LocalDate leap=LocalDate.of(2024,2,28);
        check(Core.date("今天+1天:yyyy-MM-dd",leap).equals("2024-02-29"),"leap day");
        check(Core.date("今天+2天:yyyy-MM-dd",leap).equals("2024-03-01"),"month boundary");
        check(Core.date("今天-1天:yyyy年M月d日",LocalDate.of(2026,1,1)).equals("2025年12月31日"),"year boundary");
        rejects(()->Core.date("今天+999999999999999999天",leap),"overflow");
        rejects(()->Core.date("今天:bad[",leap),"invalid date pattern");
        rejects(()->Core.tokens("{{姓名}"),"unclosed placeholder");
        rejects(()->Core.tokens("{{}}"),"empty placeholder");
        Core.Template t=new Core.Template("one","Test","{{邮箱}}","{{姓名}} / {{今天}}","{{姓名}} owes {{金额}}. {{今天+7天:yyyy-MM-dd}}");
        check(Core.variables(t).equals(new LinkedHashSet<>(Arrays.asList("邮箱","姓名","金额"))),"deduplicated variables");
        Map<String,String> v=new HashMap<>();v.put("邮箱","me@example.com");v.put("姓名","王$\\同学{{今天}}");v.put("金额","1,250.50");
        Core.Message msg=Core.generate(t,v,leap);
        check(msg.body.startsWith("王$\\同学{{今天}} owes 1,250.50"),"literal replacement, no recursive execution");
        check(msg.body.endsWith("2024-03-06"),"date snapshot");
        Map<String,String> missing=new HashMap<>(v);missing.remove("金额");rejects(()->Core.generate(t,missing,leap),"missing variable");
        rejects(()->new Core.Message("a@b.com\r\nBcc:x@y.com","subject","body"),"recipient injection");
        rejects(()->new Core.Message("a@b.com","subject\nBcc: x@y.com","body"),"subject injection");
        rejects(()->new Core.Message("broken","subject","body"),"invalid mailbox");
        rejects(()->new Core.Message("a@b.com;","subject","body"),"empty address");
        Core.Message unicode=new Core.Message("one@example.com；two@example.com","金额 & # + ?","你好\n第二行 & 100%");
        Map<String,String> params=Graph.query(unicode.mailto(true).getRawSchemeSpecificPart().split("\\?",2)[1]);
        check(params.get("subject").equals(unicode.subject),"subject URI round trip");
        check(params.get("body").equals("你好\r\n第二行 & 100%"),"body URI UTF-8 and CRLF");
        check(unicode.recipients.size()==2,"multiple recipients single message");
        String longBody=String.join("",Collections.nCopies(3000,"长"));
        Core.Message longMail=new Core.Message("one@example.com","test",longBody);
        rejects(()->longMail.mailto(true),"no long-body truncation");
        check(longMail.mailto(false).toString().length()<Core.MAILTO_LIMIT,"manual paste fallback");
        Map<String,Object> payload=Json.object(Json.parse(Json.write(unicode.payload())));
        check(Boolean.TRUE.equals(payload.get("saveToSentItems")),"save sent copy boolean");
        check(Json.str(Json.object(Json.object(payload.get("message")).get("body")),"content").equals(unicode.body),"Graph JSON Unicode");
        String odd="quote\" slash\\ control\n\t\b\u0000 中文 😀";
        check(odd.equals(Json.parse(Json.write(odd))),"JSON escapes round trip");
        rejects(()->Json.parse("{}junk"),"JSON trailing garbage");rejects(()->Json.parse("[1,]"),"JSON trailing comma");
        Path path=Files.createTempDirectory("mailtool-test-");Store store=new Store(path);
        check(store.loadTemplates().size()==1,"first-run sample");
        store.saveTemplates(Arrays.asList(t));check(store.loadTemplates().get(0).body.equals(t.body),"template persistence");
        Properties settings=new Properties();settings.setProperty("direct","true");settings.setProperty("tenant","test");store.saveSettings(settings);check(store.loadSettings().getProperty("direct").equals("true"),"settings persistence");
        store.saveTemplates(Collections.emptyList());check(store.loadTemplates().isEmpty(),"empty list stays empty");
        Files.write(path.resolve("templates.json"),"broken".getBytes(StandardCharsets.UTF_8));rejects(()->store.loadTemplates(),"corrupt data not overwritten");
        check(new String(Files.readAllBytes(path.resolve("templates.json")),StandardCharsets.UTF_8).equals("broken"),"preserve corrupt data");
        check(Graph.challenge("dBjftJeZ4CVP-mB92K27uhbUJU1p1r_wW1gFWFOEjXk").equals("E9Melhoa2OwvFrEMTJguCHaoeK1t8URWbuGJSstw-cM"),"RFC7636 PKCE vector");
        rejects(()->Graph.query("state=a&state=b"),"duplicate auth state");
        testGraph(new Store(Files.createTempDirectory("mailtool-auth-test-")),unicode);
        testHttp();
        System.out.println("PASS: "+checks+" checks. No real Microsoft login or email send was performed.");
    }
    static void testGraph(Store store,Core.Message message)throws Exception {
        AtomicInteger sends=new AtomicInteger();AtomicReference<String> expectedChallenge=new AtomicReference<>();AtomicInteger status=new AtomicInteger(202);AtomicBoolean failNetwork=new AtomicBoolean();
        Graph.Transport fake=(method,url,type,data,bearer)->{
            if(url.endsWith("/token")){
                Map<String,String> fields=Graph.query(data);
                try{check(Graph.challenge(fields.get("code_verifier")).equals(expectedChallenge.get()),"verifier matches challenge");}catch(Exception ex){throw new IOException(ex);}
                check(!fields.containsKey("client_secret"),"public client no secret");
                return new Graph.Response(200,"{\"access_token\":\"fake-access\",\"expires_in\":3600}");
            }
            if(url.contains("/me?"))return new Graph.Response(200,"{\"mail\":\"sender@example.com\"}");
            check(url.endsWith("/me/sendMail")&&method.equals("POST"),"Graph send endpoint");
            check(bearer.equals("fake-access"),"access token in transport header");
            check(Json.write(Json.parse(data)).equals(Json.write(message.payload())),"exact confirmed snapshot payload");
            sends.incrementAndGet();if(failNetwork.get())throw new IOException("simulated disconnect");
            return new Graph.Response(status.get(),status.get()==202?"":"{\"error\":{\"code\":\"test_error\"}}");
        };
        Graph graph=new Graph("11111111-1111-1111-1111-111111111111","22222222-2222-2222-2222-222222222222",store,fake);
        rejects(()->graph.send(message),"unauthenticated send blocked");
        AtomicBoolean closed=new AtomicBoolean();
        graph.ensureLogin(new Graph.LoginUi(){
            public void open(URI uri,Runnable cancel)throws Exception {
                Map<String,String> query=Graph.query(uri.getRawQuery());expectedChallenge.set(query.get("code_challenge"));
                check(query.get("code_challenge_method").equals("S256"),"S256");
                check(query.get("scope").contains("Mail.Send")&&!query.get("scope").contains("Mail.Read"),"least mail permission");
                String callback=query.get("redirect_uri");
                Graph.Http http=new Graph.Http();
                check(http.request("GET",callback+"?state=wrong&code=bad","",null,"").status==400,"invalid state refused");
                check(http.request("GET",callback+"?state="+Core.encode(query.get("state"))+"&code=fake-code","",null,"").status==200,"valid callback accepted");
            }
            public void close(){closed.set(true);}
        });
        check(closed.get(),"login callback cleaned up");check(graph.account().equals("sender@example.com"),"confirmed sender comes from Graph");check(sends.get()==0,"login does not send");
        graph.send(message);check(sends.get()==1,"one send call");
        status.set(403);rejects(()->graph.send(message),"403 surfaced");check(sends.get()==2,"403 not retried");
        status.set(500);rejects(()->graph.send(message),"500 ambiguous surfaced");check(sends.get()==3,"500 not retried");
        failNetwork.set(true);rejects(()->graph.send(message),"network ambiguity surfaced");check(sends.get()==4,"timeout not retried");
        failNetwork.set(false);status.set(401);rejects(()->graph.send(message),"expired auth surfaced");
        rejects(()->graph.send(message),"401 clears access token");check(sends.get()==5,"401 not resent");
        graph.logout();check(graph.account().isEmpty(),"logout clears sender");
    }
    static void testHttp()throws Exception {
        HttpServer server=HttpServer.create(new InetSocketAddress("127.0.0.1",0),0);AtomicInteger redirects=new AtomicInteger();
        server.createContext("/post",exchange->{
            ByteArrayOutputStream b=new ByteArrayOutputStream();Store.copy(exchange.getRequestBody(),b);
            byte[] bytes=b.toByteArray();exchange.sendResponseHeaders(200,bytes.length);exchange.getResponseBody().write(bytes);exchange.close();
        });
        server.createContext("/redirect",exchange->{exchange.getResponseHeaders().set("Location","/post");exchange.sendResponseHeaders(302,-1);exchange.close();redirects.incrementAndGet();});
        try{server.start();String base="http://127.0.0.1:"+server.getAddress().getPort();Graph.Http http=new Graph.Http();
            Graph.Response r=http.request("POST",base+"/post","application/json","{\"body\":\"中文测试\"}","");check(r.status==200&&r.body.contains("中文测试"),"real HTTP UTF-8 body");
            check(http.request("POST",base+"/redirect","application/json","{}","").status==302,"redirects not followed");check(redirects.get()==1,"single HTTP request");
        }finally{server.stop(0);}
    }
}
