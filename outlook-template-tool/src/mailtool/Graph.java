package mailtool;

import com.sun.net.httpserver.HttpServer;
import java.io.*;
import java.net.*;
import java.nio.charset.StandardCharsets;
import java.security.*;
import java.util.*;
import java.util.concurrent.*;

final class Graph {
    static final String SCOPES="https://graph.microsoft.com/Mail.Send https://graph.microsoft.com/User.Read offline_access";
    interface LoginUi {void open(URI uri,Runnable cancel)throws Exception;void close();}
    interface Transport {Response request(String method,String url,String type,String data,String bearer)throws IOException;}
    static final class Response {
        final int status;final String body;
        Response(int status,String body){this.status=status;this.body=body;}
        Map<String,Object> json(){return Json.object(Json.parse(body));}
    }
    static final class Http implements Transport {
        public Response request(String method,String url,String type,String data,String bearer)throws IOException {
            HttpURLConnection c=(HttpURLConnection)new URL(url).openConnection();
            c.setConnectTimeout(15000);c.setReadTimeout(30000);c.setInstanceFollowRedirects(false);
            c.setRequestMethod(method);c.setRequestProperty("Accept","application/json");
            if(!bearer.isEmpty())c.setRequestProperty("Authorization","Bearer "+bearer);
            try{
                if(data!=null){
                    byte[] bytes=data.getBytes(StandardCharsets.UTF_8);c.setDoOutput(true);c.setRequestProperty("Content-Type",type);
                    c.setFixedLengthStreamingMode(bytes.length);
                    try(OutputStream out=c.getOutputStream()){out.write(bytes);}
                }
                int status=c.getResponseCode();InputStream in=status>=400?c.getErrorStream():c.getInputStream();
                ByteArrayOutputStream b=new ByteArrayOutputStream();if(in!=null)try(InputStream input=in){Store.copy(input,b);}
                return new Response(status,b.toString("UTF-8"));
            }finally{c.disconnect();}
        }
    }
    final String tenant,client; final Store store;final Transport transport;
    private String access="",refresh="",account="";private long expires;
    String cacheNotice="";
    Graph(String tenant,String client,Store store)throws Exception {this(tenant,client,store,new Http());}
    Graph(String tenant,String client,Store store,Transport transport)throws Exception {
        if(!tenant.matches("[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}"))throw new IllegalArgumentException("请填写公司的目录（租户）ID，格式为 UUID");
        if(!client.matches("[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}"))throw new IllegalArgumentException("请填写应用（客户端）ID，格式为 UUID");
        this.tenant=tenant;this.client=client;this.store=store;this.transport=transport;
        try{refresh=store.readRefresh(key());}catch(Exception ex){cacheNotice="无法读取保存的登录状态，请重新登录。";}
    }
    private String key(){return tenant+"/"+client;}
    String account(){return account;}
    void ensureLogin(LoginUi ui)throws Exception {
        if(!access.isEmpty()&&System.currentTimeMillis()<expires&&!account.isEmpty())return;
        if(!refresh.isEmpty()){
            Response r=token(Json.obj("grant_type","refresh_token","refresh_token",refresh,"client_id",client,"scope",SCOPES));
            if(r.status==200){acceptToken(r);loadAccount();return;}
            String code=Json.str(r.json(),"error");
            if(!Arrays.asList("invalid_grant","interaction_required","consent_required").contains(code))throw failure(r,"更新登录状态");
            refresh="";access="";store.clearRefresh();
        }
        login(ui);loadAccount();
    }
    void logout()throws IOException {access="";refresh="";account="";expires=0;store.clearRefresh();}
    private void loadAccount()throws Exception {
        Response r=transport.request("GET","https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName","",null,access);
        if(r.status!=200){access="";throw failure(r,"读取发件账号");}
        Map<String,Object> m=r.json();String address=Json.str(m,"mail");if(address.isEmpty())address=Json.str(m,"userPrincipalName");
        if(address.isEmpty())throw new IOException("无法识别发件账号，请重新登录");
        account=address;
    }
    private void acceptToken(Response r)throws Exception {
        Map<String,Object> m=r.json();String a=Json.str(m,"access_token");
        if(a.isEmpty()||!(m.get("expires_in") instanceof Number))throw new IOException("微软登录响应缺少有效令牌");
        access=a;expires=System.currentTimeMillis()+Math.max(0,((Number)m.get("expires_in")).longValue()-120)*1000;
        String rotated=Json.str(m,"refresh_token");if(!rotated.isEmpty())refresh=rotated;
        if(!refresh.isEmpty())try{store.saveRefresh(key(),refresh);}catch(Exception ex){
            cacheNotice="登录成功，但 Windows 无法安全保存登录状态；重启后需要重新登录。";
            try{store.clearRefresh();}catch(IOException ignored){}
        }
    }
    private Response token(Map<String,Object> fields)throws IOException {
        return transport.request("POST","https://login.microsoftonline.com/"+tenant+"/oauth2/v2.0/token","application/x-www-form-urlencoded",form(fields),"");
    }
    static String form(Map<String,Object> fields){StringJoiner j=new StringJoiner("&");for(Map.Entry<String,Object> e:fields.entrySet())j.add(Core.encode(e.getKey())+"="+Core.encode(e.getValue().toString()));return j.toString();}
    static String random(){byte[] bytes=new byte[32];new SecureRandom().nextBytes(bytes);return Base64.getUrlEncoder().withoutPadding().encodeToString(bytes);}
    static String challenge(String verifier)throws Exception{return Base64.getUrlEncoder().withoutPadding().encodeToString(MessageDigest.getInstance("SHA-256").digest(verifier.getBytes(StandardCharsets.US_ASCII)));}
    static Map<String,String> query(String query)throws IOException {
        Map<String,String> m=new HashMap<>();if(query==null)return m;
        for(String p:query.split("&")){String[] pair=p.split("=",2);String k=URLDecoder.decode(pair[0],"UTF-8");if(m.containsKey(k))throw new IOException("重复的登录参数");m.put(k,pair.length==2?URLDecoder.decode(pair[1],"UTF-8"):"");}return m;
    }
    private void login(LoginUi ui)throws Exception {
        String state=random(),verifier=random();BlockingQueue<Map<String,String>> result=new ArrayBlockingQueue<>(1);
        java.util.concurrent.atomic.AtomicBoolean cancelled=new java.util.concurrent.atomic.AtomicBoolean(false);
        HttpServer server=HttpServer.create(new InetSocketAddress(InetAddress.getByName("localhost"),0),0);
        String redirect="http://localhost:"+server.getAddress().getPort()+"/";
        server.createContext("/",exchange->{
            int status=400;String response="Invalid login response. Return to the mail tool and try again.";
            try{
                Map<String,String> q=query(exchange.getRequestURI().getRawQuery());
                if(exchange.getRequestMethod().equals("GET")&&exchange.getRequestURI().getPath().equals("/")&&state.equals(q.get("state"))&&(q.containsKey("code")||q.containsKey("error"))){
                    if(result.offer(q)){status=200;response="Login response received. You can close this page and return to the mail tool.";}
                }
            }catch(Exception ignored){}
            byte[] bytes=response.getBytes(StandardCharsets.UTF_8);exchange.getResponseHeaders().set("Content-Type","text/plain; charset=utf-8");exchange.getResponseHeaders().set("Cache-Control","no-store");
            exchange.sendResponseHeaders(status,bytes.length);try(OutputStream out=exchange.getResponseBody()){out.write(bytes);}exchange.close();
        });
        try{
            server.start();
            URI auth=URI.create("https://login.microsoftonline.com/"+tenant+"/oauth2/v2.0/authorize?"+form(Json.obj("client_id",client,"response_type","code","redirect_uri",redirect,"response_mode","query","scope",SCOPES,"state",state,"code_challenge",challenge(verifier),"code_challenge_method","S256","prompt","select_account")));
            ui.open(auth,()->cancelled.set(true));
            long end=System.nanoTime()+TimeUnit.MINUTES.toNanos(5);Map<String,String> q=null;
            while(System.nanoTime()<end && !cancelled.get()){q=result.poll(1,TimeUnit.SECONDS);if(q!=null)break;}
            if(cancelled.get())throw new CancellationException("已取消登录");
            if(q==null)throw new IOException("登录等待已超时，请重新登录");
            if(q.containsKey("error"))throw new IOException("微软未完成授权："+q.get("error")+"。请检查公司授权策略。");
            Response r=token(Json.obj("client_id",client,"grant_type","authorization_code","code",q.get("code"),"redirect_uri",redirect,"code_verifier",verifier,"scope",SCOPES));
            if(cancelled.get())throw new CancellationException("已取消登录");
            if(r.status!=200)throw failure(r,"登录授权");acceptToken(r);
        }finally{server.stop(0);ui.close();}
    }
    void send(Core.Message message)throws IOException {
        if(access.isEmpty()||account.isEmpty())throw new IOException("请先登录 Microsoft 365 账号");
        final Response r;
        // Never retry this POST: a timeout may happen after Microsoft accepted the mail.
        try{r=transport.request("POST","https://graph.microsoft.com/v1.0/me/sendMail","application/json; charset=utf-8",Json.write(message.payload()),access);}
        catch(IOException ex){throw new IOException("发送结果尚不确定（网络中断或超时）。请先检查 Outlook 的已发送邮件，再决定是否重试，避免重复发送。",ex);}
        if(r.status!=202){
            if(r.status==401){access="";expires=0;}
            if(r.status>=500)throw new IOException("微软服务暂时异常，发送结果可能不确定。请先检查已发送邮件，工具没有自动重试。（HTTP "+r.status+"）");
            throw failure(r,"发送邮件");
        }
    }
    static IOException failure(Response r,String action){
        String code="";try{Map<String,Object> m=r.json();Object err=m.get("error");code=err instanceof Map?Json.str(Json.object(err),"code"):Json.str(m,"error");}catch(RuntimeException ignored){}
        return new IOException(action+"失败（HTTP "+r.status+(code.isEmpty()?"":" / "+code)+"）。"
            +(r.status==403?"请联系公司管理员检查 Mail.Send 权限和邮箱访问策略。":r.status==429?"请求过于频繁，请稍后再试。":r.status==401?"请重新登录后再次确认发送。":"请检查网络、账号和应用注册配置。"));
    }
}
