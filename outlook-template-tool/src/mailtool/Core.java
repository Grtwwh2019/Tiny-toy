package mailtool;

import java.net.*;
import java.time.*;
import java.time.format.*;
import java.util.*;
import java.util.regex.*;

final class Core {
    private static final Pattern TOKEN=Pattern.compile("\\{\\{([^{}]+)}}");
    private static final Pattern DATE=Pattern.compile("今天(?:([+-]\\d+)天)?(?::(.+))?");
    static final int MAILTO_LIMIT=1800;
    static final class Template {
        final String id,name,to,subject,body;
        Template(String id,String name,String to,String subject,String body){
            this.id=id;this.name=name;this.to=to;this.subject=subject;this.body=body;
        }
        public String toString(){return name;}
        Map<String,Object> json(){return Json.obj("id",id,"name",name,"to",to,"subject",subject,"body",body);}
        static Template from(Object o){
            Map<String,Object> m=Json.object(o);
            Template t=new Template(Json.str(m,"id"),Json.str(m,"name"),Json.str(m,"to"),Json.str(m,"subject"),Json.str(m,"body"));
            if(t.id.isEmpty()||t.name.trim().isEmpty())throw new IllegalArgumentException("模板缺少名称或编号");
            variables(t);return t;
        }
    }
    static List<String> tokens(String text){
        List<String> list=new ArrayList<>();Matcher m=TOKEN.matcher(text);int end=0;
        while(m.find()){
            String outside=text.substring(end,m.start());if(outside.contains("{{")||outside.contains("}}"))throw new IllegalArgumentException("模板占位符括号不完整");
            list.add(m.group(1).trim());end=m.end();
        }
        String tail=text.substring(end);if(tail.contains("{{")||tail.contains("}}"))throw new IllegalArgumentException("模板占位符括号不完整");
        return list;
    }
    static Set<String> variables(Template t){
        Set<String> vars=new LinkedHashSet<>();
        for(String text:Arrays.asList(t.to,t.subject,t.body))for(String token:tokens(text)){
            if(token.startsWith("今天")){date(token,LocalDate.of(2026,1,1));}
            else {if(token.isEmpty()||token.contains(":")||token.contains("\n"))throw new IllegalArgumentException("无效变量："+token);vars.add(token);}
        }
        return vars;
    }
    static String date(String token,LocalDate today){
        Matcher m=DATE.matcher(token);if(!m.matches())throw new IllegalArgumentException("日期函数格式应为 {{今天:yyyy-MM-dd}} 或 {{今天+7天:yyyy-MM-dd}}");
        try{
            long delta=m.group(1)==null?0:Long.parseLong(m.group(1));
            if(Math.abs(delta)>365000 || delta==Long.MIN_VALUE)throw new IllegalArgumentException();
            String pattern=m.group(2)==null?"yyyy-MM-dd":m.group(2);
            return today.plusDays(delta).format(DateTimeFormatter.ofPattern(pattern,Locale.SIMPLIFIED_CHINESE));
        }catch(RuntimeException ex){throw new IllegalArgumentException("无效日期格式或偏移量："+token);}
    }
    static String render(String text,Map<String,String> vars,LocalDate today){
        tokens(text); Matcher m=TOKEN.matcher(text);StringBuffer b=new StringBuffer();
        while(m.find()){
            String key=m.group(1).trim(),v;
            if(key.startsWith("今天"))v=date(key,today);
            else {v=vars.get(key);if(v==null||v.trim().isEmpty())throw new IllegalArgumentException("请填写变量："+key);}
            m.appendReplacement(b,Matcher.quoteReplacement(v));
        }m.appendTail(b);return b.toString();
    }
    static Message generate(Template t,Map<String,String> vars,LocalDate today){
        return new Message(render(t.to,vars,today),render(t.subject,vars,today),render(t.body,vars,today));
    }
    static final class Message {
        final List<String> recipients;final String subject,body;
        Message(String to,String subject,String body){
            if(to.indexOf('\n')>=0||to.indexOf('\r')>=0)throw new IllegalArgumentException("收件地址不能包含换行");
            List<String> addresses=new ArrayList<>();
            for(String piece:to.split("[;,；，]",-1)){
                String a=piece.trim();if(a.isEmpty())throw new IllegalArgumentException("请填写有效收件地址，用分号分隔多个地址");
                if(!a.matches("[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9](?:[A-Za-z0-9.-]*[A-Za-z0-9])?")||a.contains(".."))throw new IllegalArgumentException("收件地址格式错误："+a);
                if(!addresses.contains(a))addresses.add(a);
            }
            if(subject.trim().isEmpty())throw new IllegalArgumentException("请填写邮件主题");
            if(subject.indexOf('\n')>=0||subject.indexOf('\r')>=0)throw new IllegalArgumentException("主题不能包含换行");
            if(body.trim().isEmpty())throw new IllegalArgumentException("请填写邮件正文");
            this.recipients=Collections.unmodifiableList(addresses);this.subject=subject;this.body=body;
        }
        String preview(String from){return (from.isEmpty()?"":"发件人："+from+"\n")+"收件人："+String.join("; ",recipients)+"\n主题："+subject+"\n\n"+body;}
        URI mailto(boolean includeBody){
            StringJoiner to=new StringJoiner(",");for(String a:recipients)to.add(encode(a));
            String uri="mailto:"+to+"?subject="+encode(subject)+(includeBody?"&body="+encode(body.replace("\r\n","\n").replace("\r","\n").replace("\n","\r\n")):"");
            if(uri.length()>MAILTO_LIMIT)throw new IllegalArgumentException("邮件链接过长，请使用复制正文方式，或缩短收件人和主题");
            return URI.create(uri);
        }
        Map<String,Object> payload(){
            List<Object> to=new ArrayList<>();for(String a:recipients)to.add(Json.obj("emailAddress",Json.obj("address",a)));
            return Json.obj("message",Json.obj("subject",subject,"body",Json.obj("contentType","Text","content",body),"toRecipients",to),"saveToSentItems",true);
        }
    }
    static String encode(String s){try{return URLEncoder.encode(s,"UTF-8").replace("+","%20");}catch(Exception ex){throw new IllegalStateException(ex);}}
}
