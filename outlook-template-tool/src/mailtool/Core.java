package mailtool;

import java.net.*;
import java.time.*;
import java.time.format.*;
import java.util.*;
import java.util.regex.*;

final class Core {
    private static final Pattern TOKEN=Pattern.compile("\\{\\{([^{}]+)}}");
    private static final Pattern DATE=Pattern.compile("今天(?:([+-]\\d+)天)?(?::(.+))?");
    private static final Pattern NAMED_ADDRESS=Pattern.compile("^(.*?)\\s*<\\s*([^<>]+)\\s*>$");
    private static final Pattern EMAIL=Pattern.compile("[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9](?:[A-Za-z0-9.-]*[A-Za-z0-9])?");
    static final int MAILTO_LIMIT=1800;

    static final class Template {
        final String id,name,to,cc,subject,body;
        Template(String id,String name,String to,String cc,String subject,String body){
            this.id=id;this.name=name;this.to=to;this.cc=cc;this.subject=subject;this.body=body;
        }
        public String toString(){return name;}
        Map<String,Object> json(){return Json.obj("id",id,"name",name,"to",to,"cc",cc,"subject",subject,"body",body);}
        static Template from(Object o){
            Map<String,Object> m=Json.object(o);
            Template t=new Template(Json.str(m,"id"),Json.str(m,"name"),Json.str(m,"to"),Json.str(m,"cc"),Json.str(m,"subject"),Json.str(m,"body"));
            if(t.id.isEmpty()||t.name.trim().isEmpty())throw new IllegalArgumentException("模板缺少名称或编号");
            variables(t);return t;
        }
    }

    static final class Address {
        final String name,email;
        Address(String name,String email){this.name=name;this.email=email;}
        String display(){return name.isEmpty()?email:name+" <"+email+">";}
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
        for(String text:Arrays.asList(t.to,t.cc,t.subject,t.body))for(String token:tokens(text)){
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
        tokens(text);Matcher m=TOKEN.matcher(text);StringBuffer b=new StringBuffer();
        while(m.find()){
            String key=m.group(1).trim(),v;
            if(key.startsWith("今天"))v=date(key,today);
            else {v=vars.get(key);if(v==null||v.trim().isEmpty())throw new IllegalArgumentException("请填写变量："+key);}
            m.appendReplacement(b,Matcher.quoteReplacement(v));
        }m.appendTail(b);return b.toString();
    }

    static Message generate(Template t,Map<String,String> vars,LocalDate today){
        return new Message(render(t.to,vars,today),render(t.cc,vars,today),render(t.subject,vars,today),render(t.body,vars,today));
    }

    static List<Address> addresses(String value,boolean required,String label){
        if(value.indexOf('\n')>=0||value.indexOf('\r')>=0)throw new IllegalArgumentException(label+"不能包含换行");
        if(value.trim().isEmpty()){
            if(required)throw new IllegalArgumentException("请填写"+label);
            return Collections.emptyList();
        }
        List<Address> result=new ArrayList<>();Set<String> seen=new HashSet<>();
        for(String piece:value.split("[;；]",-1)){
            String raw=piece.trim();if(raw.isEmpty())throw new IllegalArgumentException(label+"中存在空地址，请用分号正确分隔");
            String name="",email=raw;Matcher named=NAMED_ADDRESS.matcher(raw);
            if(named.matches()){
                name=named.group(1).trim();email=named.group(2).trim();
                if(name.isEmpty())throw new IllegalArgumentException(label+"显示名称不能为空："+raw);
            }else if(raw.indexOf('<')>=0||raw.indexOf('>')>=0){throw new IllegalArgumentException(label+"格式错误："+raw);}
            if(name.indexOf('<')>=0||name.indexOf('>')>=0||name.indexOf(';')>=0||name.indexOf('；')>=0)throw new IllegalArgumentException(label+"显示名称格式错误："+raw);
            if(!EMAIL.matcher(email).matches()||email.contains(".."))throw new IllegalArgumentException(label+"邮箱格式错误："+email);
            String key=email.toLowerCase(Locale.ROOT);if(seen.add(key))result.add(new Address(name,email));
        }
        return Collections.unmodifiableList(result);
    }

    static final class Message {
        final List<Address> recipients,cc;final String subject,body;
        Message(String to,String cc,String subject,String body){
            this.recipients=addresses(to,true,"收件人");this.cc=addresses(cc,false,"抄送人");
            if(subject.trim().isEmpty())throw new IllegalArgumentException("请填写邮件主题");
            if(subject.indexOf('\n')>=0||subject.indexOf('\r')>=0)throw new IllegalArgumentException("主题不能包含换行");
            if(body.trim().isEmpty())throw new IllegalArgumentException("请填写邮件正文");
            this.subject=subject;this.body=body;
        }
        String preview(String from){
            StringBuilder b=new StringBuilder();if(!from.isEmpty())b.append("发件人：").append(from).append('\n');
            b.append("收件人：").append(joinDisplay(recipients));if(!cc.isEmpty())b.append("\n抄送人：").append(joinDisplay(cc));
            return b.append("\n主题：").append(subject).append("\n\n").append(body).toString();
        }
        URI mailto(boolean includeBody){
            StringJoiner to=new StringJoiner(",");for(Address a:recipients)to.add(encode(a.email));
            StringBuilder uri=new StringBuilder("mailto:").append(to).append("?subject=").append(encode(subject));
            if(!cc.isEmpty()){StringJoiner copies=new StringJoiner(",");for(Address a:cc)copies.add(a.email);uri.append("&cc=").append(encode(copies.toString()));}
            if(includeBody)uri.append("&body=").append(encode(body.replace("\r\n","\n").replace("\r","\n").replace("\n","\r\n")));
            if(uri.length()>MAILTO_LIMIT)throw new IllegalArgumentException("邮件链接过长，请使用复制正文方式，或缩短收件人、抄送人和主题");
            return URI.create(uri.toString());
        }
        Map<String,Object> outlookPayload(String operation){
            List<Object> to=new ArrayList<>(),copies=new ArrayList<>();
            for(Address a:recipients)to.add(a.display());for(Address a:cc)copies.add(a.display());
            return Json.obj("operation",operation,"to",to,"cc",copies,"subject",subject,"body",body);
        }
    }

    static String joinDisplay(List<Address> addresses){StringJoiner j=new StringJoiner("; ");for(Address a:addresses)j.add(a.display());return j.toString();}
    static String encode(String s){try{return URLEncoder.encode(s,"UTF-8").replace("+","%20");}catch(Exception ex){throw new IllegalStateException(ex);}}
}
