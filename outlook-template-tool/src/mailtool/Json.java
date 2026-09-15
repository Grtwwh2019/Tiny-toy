package mailtool;

import java.util.*;

/** Small JSON codec for the fixed Graph protocol and local template documents. */
final class Json {
    static Map<String,Object> obj(Object... entries) {
        Map<String,Object> m = new LinkedHashMap<>();
        for (int i=0;i<entries.length;i+=2) m.put((String)entries[i], entries[i+1]);
        return m;
    }
    @SuppressWarnings("unchecked") static Map<String,Object> object(Object value) {
        if (!(value instanceof Map)) throw new IllegalArgumentException("JSON 对象格式错误");
        return (Map<String,Object>)value;
    }
    static String str(Map<String,Object> m, String key) {
        Object v=m.get(key); return v instanceof String ? (String)v : "";
    }
    static String write(Object value) {
        if (value==null) return "null";
        if (value instanceof String) {
            StringBuilder b=new StringBuilder("\"");
            for (char c:((String)value).toCharArray()) {
                switch(c) {
                    case '"': b.append("\\\""); break;
                    case '\\': b.append("\\\\"); break;
                    case '\n': b.append("\\n"); break;
                    case '\r': b.append("\\r"); break;
                    case '\t': b.append("\\t"); break;
                    default: if(c<32) b.append(String.format("\\u%04x",(int)c)); else b.append(c);
                }
            }
            return b.append('"').toString();
        }
        if(value instanceof Boolean || value instanceof Number) return value.toString();
        StringJoiner j=new StringJoiner(",",value instanceof Map ? "{" : "[",value instanceof Map ? "}" : "]");
        if(value instanceof Map) {
            for(Map.Entry<?,?> e:((Map<?,?>)value).entrySet()) j.add(write(e.getKey().toString())+":"+write(e.getValue()));
        } else if(value instanceof Iterable) {
            for(Object v:(Iterable<?>)value) j.add(write(v));
        } else throw new IllegalArgumentException("Unsupported JSON value");
        return j.toString();
    }
    static Object parse(String text) {
        Parser p=new Parser(text); Object v=p.value(0); p.ws();
        if(p.i!=text.length()) throw p.bad(); return v;
    }
    private static final class Parser {
        final String s; int i;
        Parser(String s){this.s=s;}
        IllegalArgumentException bad(){return new IllegalArgumentException("JSON 格式错误，位置 "+i);}
        void ws(){while(i<s.length() && " \r\n\t".indexOf(s.charAt(i))>=0)i++;}
        boolean eat(char c){ws();if(i<s.length() && s.charAt(i)==c){i++;return true;}return false;}
        Object value(int depth){
            if(depth>64) throw bad(); ws(); if(i>=s.length())throw bad(); char c=s.charAt(i);
            if(c=='"')return string();
            if(eat('{')){
                Map<String,Object> m=new LinkedHashMap<>(); if(eat('}'))return m;
                do {ws(); if(i>=s.length()||s.charAt(i)!='"')throw bad();String k=string();if(!eat(':'))throw bad();m.put(k,value(depth+1));}while(eat(','));
                if(!eat('}'))throw bad();return m;
            }
            if(eat('[')){
                List<Object> a=new ArrayList<>();if(eat(']'))return a;
                do{a.add(value(depth+1));}while(eat(','));if(!eat(']'))throw bad();return a;
            }
            for(String lit:Arrays.asList("true","false","null"))if(s.startsWith(lit,i)){
                i+=lit.length();return lit.equals("null")?null:Boolean.valueOf(lit);
            }
            int start=i;while(i<s.length() && "-+0123456789.eE".indexOf(s.charAt(i))>=0)i++;
            String n=s.substring(start,i);
            if(!n.matches("-?(0|[1-9][0-9]*)(\\.[0-9]+)?([eE][+-]?[0-9]+)?"))throw bad();
            try{return new java.math.BigDecimal(n);}catch(NumberFormatException ex){throw bad();}
        }
        String string(){
            if(s.charAt(i++)!='"')throw bad();StringBuilder b=new StringBuilder();
            while(i<s.length()){
                char c=s.charAt(i++);if(c=='"')return b.toString();if(c<32)throw bad();
                if(c=='\\'){
                    if(i>=s.length())throw bad();char e=s.charAt(i++);
                    switch(e){
                        case '"':case '\\':case '/': b.append(e);break;
                        case 'b':b.append('\b');break;case 'f':b.append('\f');break;
                        case 'n':b.append('\n');break;case 'r':b.append('\r');break;case 't':b.append('\t');break;
                        case 'u':if(i+4>s.length())throw bad();try{b.append((char)Integer.parseInt(s.substring(i,i+4),16));}catch(NumberFormatException x){throw bad();}i+=4;break;
                        default:throw bad();
                    }
                }else b.append(c);
            }throw bad();
        }
    }
}
