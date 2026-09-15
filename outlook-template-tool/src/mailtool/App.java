package mailtool;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import javax.swing.event.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.datatransfer.StringSelection;
import java.awt.event.*;
import java.net.URI;
import java.nio.channels.*;
import java.nio.file.*;
import java.time.LocalDate;
import java.util.*;
import java.util.List;
import java.util.concurrent.*;

public final class App extends JFrame {
    private static final long serialVersionUID=1L;
    private final Store store;
    private final Properties settings;
    private final Outlook outlook;
    private List<Core.Template> templates;
    private final JComboBox<Core.Template> choice=new JComboBox<>();
    private final DefaultTableModel values=new DefaultTableModel(new Object[]{"变量","填写内容"},0){private static final long serialVersionUID=1L;public boolean isCellEditable(int r,int c){return c==1;}};
    private final JTable table=new JTable(values);
    private final JTextArea preview=area(false,12);
    private final JLabel status=new JLabel("就绪：选模板，填变量，再打开 Outlook 或确认发送。"),outlookLabel=new JLabel("经典 Outlook 状态：尚未检测"),modeLabel=new JLabel();
    private final JCheckBox direct=new JCheckBox("确认后直接发送（仅经典 Outlook，不打开编辑窗口）");
    private final JButton send=new JButton(),checkOutlook=new JButton("检测经典 Outlook");
    private final DefaultListModel<Core.Template> templateModel=new DefaultListModel<>();
    private final JList<Core.Template> templateList=new JList<>(templateModel);
    private final JTextField name=new JTextField(),to=new JTextField(),cc=new JTextField(),subject=new JTextField();
    private final JTextArea body=area(true,12);
    private final List<JButton> busyButtons=new ArrayList<>();
    private String editId=null;
    private boolean busy=false,loadingEditor=false,dirty=false,reverting=false;
    private static FileChannel lockChannel;
    private static FileLock processLock;

    private App(Store store)throws Exception {this(store,new Outlook());}
    private App(Store store,Outlook outlook)throws Exception {
        super("Outlook 模板邮件工具");this.store=store;this.outlook=outlook;
        Properties previous=store.loadSettings();settings=new Properties();settings.setProperty("direct",previous.getProperty("direct","false"));
        store.saveSettings(settings);store.clearLegacyCredentials();templates=store.loadTemplates();
        setDefaultCloseOperation(DO_NOTHING_ON_CLOSE);
        addWindowListener(new WindowAdapter(){public void windowClosing(WindowEvent e){
            if(busy){info("操作仍在进行，请稍候。");return;}
            if(discardEditor()){dispose();System.exit(0);}
        }});
        JPanel root=new JPanel(new BorderLayout(0,12));root.setBorder(new EmptyBorder(20,24,14,24));setContentPane(root);
        JLabel heading=new JLabel("把重复邮件，交给模板");heading.setFont(heading.getFont().deriveFont(Font.BOLD,25f));
        JPanel header=new JPanel(new BorderLayout());header.add(heading,BorderLayout.NORTH);header.add(new JLabel("填写一次变量 · 预览完整内容 · 使用桌面 Outlook 当前账号"),BorderLayout.SOUTH);root.add(header,BorderLayout.NORTH);
        JTabbedPane tabs=new JTabbedPane();tabs.addTab("  生成邮件  ",composePanel());tabs.addTab("  编辑模板  ",templatePanel());tabs.addTab("  Outlook 与设置  ",settingsPanel());root.add(tabs,BorderLayout.CENTER);
        status.setBorder(new EmptyBorder(8,0,0,0));root.add(status,BorderLayout.SOUTH);
        direct.setSelected(Boolean.parseBoolean(settings.getProperty("direct","false")));updateMode();refreshTemplates(templates.isEmpty()?null:templates.get(0).id);
        setMinimumSize(new Dimension(850,680));setSize(1050,820);setLocationRelativeTo(null);
    }

    private JPanel composePanel(){
        JPanel panel=panel();JPanel top=new JPanel(new BorderLayout(12,8));top.add(new JLabel("选择模板"),BorderLayout.WEST);top.add(choice,BorderLayout.CENTER);panel.add(top,BorderLayout.NORTH);
        choice.addActionListener(e->fillVariables());table.setRowHeight(32);table.putClientProperty("terminateEditOnFocusLost",true);table.getColumnModel().getColumn(0).setPreferredWidth(100);table.getColumnModel().getColumn(1).setPreferredWidth(230);
        JPanel left=new JPanel(new BorderLayout(0,8));left.add(new JLabel("填写变量（日期在生成时自动计算）"),BorderLayout.NORTH);left.add(new JScrollPane(table),BorderLayout.CENTER);
        JButton update=new JButton("更新预览");update.addActionListener(e->showPreview(true));left.add(update,BorderLayout.SOUTH);
        JPanel right=new JPanel(new BorderLayout(0,8));right.add(new JLabel("邮件预览"),BorderLayout.NORTH);right.add(new JScrollPane(preview),BorderLayout.CENTER);
        JSplitPane split=new JSplitPane(JSplitPane.HORIZONTAL_SPLIT,left,right);split.setResizeWeight(.35);split.setDividerLocation(320);split.setBorder(null);panel.add(split,BorderLayout.CENTER);
        JPanel footer=new JPanel(new BorderLayout());footer.add(modeLabel,BorderLayout.CENTER);send.addActionListener(e->sendAction());busyButtons.add(send);footer.add(send,BorderLayout.EAST);panel.add(footer,BorderLayout.SOUTH);
        values.addTableModelListener(e->showPreview(false));return panel;
    }

    private JPanel templatePanel(){
        JPanel panel=panel();JPanel left=new JPanel(new BorderLayout(0,8));left.setPreferredSize(new Dimension(200,400));left.add(new JLabel("我的模板"),BorderLayout.NORTH);left.add(new JScrollPane(templateList),BorderLayout.CENTER);
        templateList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        templateList.addListSelectionListener(e->{if(!e.getValueIsAdjusting()&&!reverting){
            Core.Template t=templateList.getSelectedValue();if(t!=null&&!Objects.equals(t.id,editId)){
                if(!discardEditor()){reverting=true;selectEditor(editId);reverting=false;return;}loadEditor(t);
            }
        }});
        JButton add=new JButton("新建"),delete=new JButton("删除");JPanel tools=new JPanel(new FlowLayout(FlowLayout.LEFT,4,0));tools.add(add);tools.add(delete);left.add(tools,BorderLayout.SOUTH);
        add.addActionListener(e->{if(discardEditor()){reverting=true;templateList.clearSelection();reverting=false;loadEditor(null);name.requestFocusInWindow();}});
        delete.addActionListener(e->{
            if(editId==null)return;
            if(JOptionPane.showConfirmDialog(this,"删除模板“"+name.getText()+"”？","删除模板",JOptionPane.OK_CANCEL_OPTION)!=JOptionPane.OK_OPTION)return;
            List<Core.Template> next=new ArrayList<>(templates);next.removeIf(t->t.id.equals(editId));
            try{store.saveTemplates(next);templates=next;dirty=false;editId=null;refreshTemplates(next.isEmpty()?null:next.get(0).id);}catch(Exception ex){error(ex);}
        });
        JPanel editor=new JPanel(new BorderLayout(0,10));JPanel fields=new JPanel(new GridLayout(0,1,0,6));fields.add(field("模板名称",name));fields.add(field("收件人（多人用分号分隔）",to));fields.add(field("抄送人（可选，多人用分号分隔）",cc));fields.add(field("邮件主题",subject));editor.add(fields,BorderLayout.NORTH);
        JPanel content=new JPanel(new BorderLayout(0,8));content.add(new JLabel("正文（纯文本）"),BorderLayout.NORTH);content.add(new JScrollPane(body),BorderLayout.CENTER);editor.add(content,BorderLayout.CENTER);
        JPanel bottom=new JPanel(new BorderLayout(0,8));JTextArea hint=area(false,4);hint.setText("地址格式：Jim zheng<jim_zheng@mail.com>；多人用英文或中文分号分隔。\n手动变量：{{姓名}}、{{金额}}、{{收件邮箱}}\n日期函数：{{今天:yyyy-MM-dd}}、{{今天+7天:yyyy-MM-dd}}、{{今天-1天:yyyy年M月d日}}\n收件人和抄送人也可使用变量；抄送人可以留空。");hint.setBackground(editor.getBackground());bottom.add(hint,BorderLayout.CENTER);
        JButton save=new JButton("保存模板");save.addActionListener(e->saveTemplate());bottom.add(save,BorderLayout.SOUTH);editor.add(bottom,BorderLayout.SOUTH);
        JSplitPane split=new JSplitPane(JSplitPane.HORIZONTAL_SPLIT,left,editor);split.setDividerLocation(220);split.setBorder(null);panel.add(split,BorderLayout.CENTER);
        DocumentListener listener=new DocumentListener(){public void insertUpdate(DocumentEvent e){mark();}public void removeUpdate(DocumentEvent e){mark();}public void changedUpdate(DocumentEvent e){mark();}void mark(){if(!loadingEditor)dirty=true;}};
        for(JTextField f:Arrays.asList(name,to,cc,subject))f.getDocument().addDocumentListener(listener);body.getDocument().addDocumentListener(listener);
        busyButtons.add(add);busyButtons.add(delete);busyButtons.add(save);return panel;
    }

    private JPanel settingsPanel(){
        JPanel panel=panel();JPanel stack=new JPanel();stack.setLayout(new BoxLayout(stack,BoxLayout.Y_AXIS));
        JTextArea intro=area(false,6);intro.setText("工具不再要求 Microsoft 365 登录，也不保存授权令牌。\n\n经典 Outlook：直接使用桌面 Outlook 当前配置的账号，可打开邮件，也可在确认后发送。\n新版 Outlook：可通过 Windows 默认邮件应用打开编辑窗口；不支持从本工具直接发送。\n\n如果同时安装了两个版本，“打开邮件”会优先调用经典 Outlook。");intro.setBackground(panel.getBackground());stack.add(intro);stack.add(Box.createVerticalStrut(16));
        direct.addActionListener(e->{try{persistSettings();updateMode();status.setText("发送模式已保存。");}catch(Exception ex){error(ex);}});stack.add(direct);stack.add(Box.createVerticalStrut(16));
        checkOutlook.addActionListener(e->runWork(()->{String account=outlook.account();SwingUtilities.invokeLater(()->outlookLabel.setText("经典 Outlook 可用 · 当前账号："+account));return "已检测到经典 Outlook，当前账号："+account;}));
        busyButtons.add(checkOutlook);JPanel actions=new JPanel(new FlowLayout(FlowLayout.LEFT,0,0));actions.add(checkOutlook);stack.add(actions);stack.add(Box.createVerticalStrut(14));stack.add(outlookLabel);
        JTextArea note=area(false,5);note.setText("\n直接发送前仍会显示发件账号、收件人、抄送人、主题和完整正文。\nOutlook 或公司的安全策略可能另外显示程序访问提示。\n直接发送不会自动添加 Outlook 签名，请把署名写入模板。\n若检测不到经典 Outlook，请关闭直接发送模式，使用“打开 Outlook”。");note.setBackground(panel.getBackground());stack.add(note);
        stack.setAlignmentX(Component.LEFT_ALIGNMENT);panel.add(stack,BorderLayout.NORTH);return panel;
    }

    private void persistSettings()throws Exception {
        Properties next=new Properties();next.setProperty("direct",Boolean.toString(direct.isSelected()));store.saveSettings(next);settings.clear();settings.putAll(next);
    }
    private void updateMode(){boolean enabled=direct.isSelected();modeLabel.setText(enabled?"模式：经典 Outlook 确认后直接发送":"模式：打开 Outlook 编辑邮件");send.setText(enabled?"发送…":"打开 Outlook");}
    private void fillVariables(){Map<String,String> old=readValues();values.setRowCount(0);Core.Template t=(Core.Template)choice.getSelectedItem();if(t!=null)for(String key:Core.variables(t))values.addRow(new Object[]{key,old.getOrDefault(key,"")});showPreview(false);}
    private Map<String,String> readValues(){Map<String,String> result=new LinkedHashMap<>();for(int i=0;i<values.getRowCount();i++)result.put(values.getValueAt(i,0).toString(),Objects.toString(values.getValueAt(i,1),""));return result;}
    private Core.Message generate(){if(table.isEditing())table.getCellEditor().stopCellEditing();Core.Template t=(Core.Template)choice.getSelectedItem();if(t==null)throw new IllegalArgumentException("请先新建并保存一个模板");return Core.generate(t,readValues(),LocalDate.now());}
    private void showPreview(boolean explicit){
        try{Core.Template t=(Core.Template)choice.getSelectedItem();if(t==null){preview.setText("请在“编辑模板”中创建模板。");return;}if(explicit&&table.isEditing())table.getCellEditor().stopCellEditing();Core.Message m=Core.generate(t,readValues(),LocalDate.now());preview.setText(m.preview(""));preview.setCaretPosition(0);}
        catch(IllegalArgumentException ex){preview.setText("待填写或修正：\n"+ex.getMessage());if(explicit)error(ex);}
    }

    private void sendAction(){
        if(busy)return;final Core.Message message;
        try{message=generate();preview.setText(message.preview(""));}catch(Exception ex){error(ex);return;}
        if(!direct.isSelected()){runWork(()->openMessage(message));return;}
        runWork(()->{
            final String account;
            try{account=outlook.account();}catch(Outlook.NotAvailableException ex){throw new IllegalStateException("直接发送仅支持经典 Outlook。当前未检测到经典 Outlook 本机接口；请关闭直接发送模式，改用“打开 Outlook”。\n\n"+ex.getMessage());}
            final boolean[] confirmed={false};SwingUtilities.invokeAndWait(()->confirmed[0]=confirmMail(message,account));
            if(!confirmed[0])return "已取消发送，邮件未提交。";
            outlook.send(message);return "已通过经典 Outlook 提交发送，请在“已发送邮件”中确认。";
        });
    }

    private String openMessage(Core.Message message)throws Exception {
        try{String account=outlook.open(message);return "已通过经典 Outlook 打开邮件窗口，发件账号："+account;}
        catch(Outlook.NotAvailableException unavailable){
            URI uri;
            try{uri=message.mailto(true);}catch(IllegalArgumentException tooLong){
                final int[] answer={JOptionPane.CANCEL_OPTION};
                SwingUtilities.invokeAndWait(()->answer[0]=JOptionPane.showConfirmDialog(this,"未检测到经典 Outlook，且正文较长。\n是否复制完整正文并用默认邮件应用打开？打开后请在正文处按 Ctrl+V。","长邮件",JOptionPane.OK_CANCEL_OPTION));
                if(answer[0]!=JOptionPane.OK_OPTION)return "已取消打开邮件。";
                uri=message.mailto(false);SwingUtilities.invokeAndWait(()->Toolkit.getDefaultToolkit().getSystemClipboard().setContents(new StringSelection(message.body),null));
            }
            if(!Desktop.isDesktopSupported()||!Desktop.getDesktop().isSupported(Desktop.Action.MAIL))throw new IllegalStateException("系统未提供邮件打开功能，请将 Outlook 设为 Windows 默认邮件应用。");
            Desktop.getDesktop().mail(uri);return "未检测到经典 Outlook，已请求用 Windows 默认邮件应用打开。请检查收件人、抄送人和正文。";
        }
    }

    private boolean confirmMail(Core.Message message,String from){
        JTextArea full=area(false,18);full.setText(message.preview(from));full.setCaretPosition(0);JScrollPane scroll=new JScrollPane(full);scroll.setPreferredSize(new Dimension(660,420));
        JPanel panel=new JPanel(new BorderLayout(0,10));panel.add(new JLabel("确认后立即通过经典 Outlook 提交发送，不会打开编辑窗口。"),BorderLayout.NORTH);panel.add(scroll,BorderLayout.CENTER);
        Object[] options={"取消","确认发送"};return JOptionPane.showOptionDialog(this,panel,"确认发送这一封邮件",JOptionPane.DEFAULT_OPTION,JOptionPane.PLAIN_MESSAGE,null,options,options[0])==1;
    }

    private void runWork(Callable<String> action){
        if(busy)return;setBusy(true);status.setText("正在处理，请稍候…");
        new SwingWorker<String,Void>(){
            protected String doInBackground()throws Exception{return action.call();}
            protected void done(){setBusy(false);try{String result=get();status.setText(result);if(result.startsWith("已通过经典 Outlook 提交"))info(result);}catch(Exception ex){Throwable cause=ex instanceof ExecutionException?ex.getCause():ex;status.setText("操作未完成，请查看提示。");error(cause);}}
        }.execute();
    }
    private void setBusy(boolean value){busy=value;for(JButton b:busyButtons)b.setEnabled(!value);direct.setEnabled(!value);}
    private boolean discardEditor(){return !dirty||JOptionPane.showConfirmDialog(this,"模板有未保存的修改，是否放弃这些修改？","未保存的模板",JOptionPane.OK_CANCEL_OPTION)==JOptionPane.OK_OPTION;}
    private void loadEditor(Core.Template t){loadingEditor=true;editId=t==null?null:t.id;name.setText(t==null?"":t.name);to.setText(t==null?"{{收件人姓名}}<{{收件邮箱}}>":t.to);cc.setText(t==null?"":t.cc);subject.setText(t==null?"":t.subject);body.setText(t==null?"":t.body);loadingEditor=false;dirty=false;}
    private void selectEditor(String id){for(int i=0;i<templateModel.size();i++)if(templateModel.get(i).id.equals(id)){templateList.setSelectedIndex(i);return;}templateList.clearSelection();}
    private void refreshTemplates(String selected){
        reverting=true;templateModel.clear();choice.removeAllItems();for(Core.Template t:templates){templateModel.addElement(t);choice.addItem(t);}Core.Template target=null;for(Core.Template t:templates)if(t.id.equals(selected))target=t;if(target==null&&!templates.isEmpty())target=templates.get(0);if(target!=null){choice.setSelectedItem(target);templateList.setSelectedValue(target,true);}loadEditor(target);reverting=false;fillVariables();
    }
    private void saveTemplate(){
        try{
            String n=name.getText().trim();if(n.isEmpty())throw new IllegalArgumentException("请填写模板名称");if(to.getText().trim().isEmpty()||subject.getText().trim().isEmpty()||body.getText().trim().isEmpty())throw new IllegalArgumentException("收件人、主题和正文不能为空");
            String id=editId==null?UUID.randomUUID().toString():editId;Core.Template t=new Core.Template(id,n,to.getText(),cc.getText(),subject.getText(),body.getText());Core.variables(t);
            for(Core.Template other:templates)if(!other.id.equals(id)&&other.name.equals(n))throw new IllegalArgumentException("已有同名模板，请换一个名称");
            List<Core.Template> next=new ArrayList<>(templates);int index=-1;for(int i=0;i<next.size();i++)if(next.get(i).id.equals(id))index=i;if(index<0)next.add(t);else next.set(index,t);store.saveTemplates(next);templates=next;dirty=false;refreshTemplates(id);status.setText("模板已保存："+n);
        }catch(Exception ex){error(ex);}
    }
    private static JPanel panel(){JPanel p=new JPanel(new BorderLayout(12,14));p.setBorder(new EmptyBorder(16,12,14,12));return p;}
    private static JPanel field(String label,JTextField f){JPanel p=new JPanel(new BorderLayout(0,4));p.add(new JLabel(label),BorderLayout.NORTH);p.add(f,BorderLayout.CENTER);p.setMaximumSize(new Dimension(Integer.MAX_VALUE,65));return p;}
    private static JTextArea area(boolean editable,int rows){JTextArea a=new JTextArea(rows,30);a.setEditable(editable);a.setLineWrap(true);a.setWrapStyleWord(true);a.setMargin(new Insets(10,10,10,10));return a;}
    private void info(String text){JOptionPane.showMessageDialog(this,text,"提示",JOptionPane.INFORMATION_MESSAGE);}
    private void error(Throwable ex){JTextArea text=area(false,6);text.setText(ex.getMessage()==null?"操作失败，请检查设置。":ex.getMessage());JOptionPane.showMessageDialog(this,new JScrollPane(text),"操作未完成",JOptionPane.ERROR_MESSAGE);}

    public static void main(String[] args){
        SwingUtilities.invokeLater(()->{
            try{
                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());Font font=new Font("Dialog",Font.PLAIN,14);Enumeration<Object> keys=UIManager.getDefaults().keys();while(keys.hasMoreElements()){Object k=keys.nextElement();if(UIManager.get(k) instanceof Font)UIManager.put(k,font);}
                Path data=Store.defaultDir();Store store=new Store(data);lockChannel=FileChannel.open(data.resolve("app.lock"),StandardOpenOption.CREATE,StandardOpenOption.WRITE);processLock=lockChannel.tryLock();
                if(processLock==null){JOptionPane.showMessageDialog(null,"工具已经打开，请使用现有窗口。");lockChannel.close();return;}new App(store).setVisible(true);
            }catch(Exception ex){JOptionPane.showMessageDialog(null,"无法启动："+ex.getMessage()+"\n已有数据未被覆盖。","启动失败",JOptionPane.ERROR_MESSAGE);}
        });
    }
}
