package cn.com.taiji.converter.jacob;

import cn.com.taiji.converter.vo.ActiveXComVo;
import cn.com.taiji.converter.vo.ConvertedTarget;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import lombok.extern.slf4j.Slf4j;
import org.springframework.util.FileCopyUtils;
import org.springframework.util.ObjectUtils;
import org.springframework.util.StringUtils;
import sun.misc.BASE64Decoder;

import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;
import java.nio.charset.Charset;
import java.util.*;
import java.util.concurrent.*;

@Slf4j
public class JacobOffice {

    /**
     * 核心线程数
     */
    private static int processorNum = Runtime.getRuntime().availableProcessors();

    /**
     * 转换线程池
     */
    private static ExecutorService executor = new ThreadPoolExecutor(processorNum * 2, processorNum * 2, 0L, TimeUnit.MILLISECONDS, new LinkedBlockingQueue<>());

    /**
     * 转换队列
     */
    private static List<ActiveXComVo> wordApps = Collections.synchronizedList(new ArrayList<>());

    /**
     * 转换队列
     */
    public static BlockingQueue<ConvertedTarget> targets = new LinkedBlockingQueue<>();

    /**
     * word pdf 临时目录
     */
    private static final String TMP_DIR = "C:\\wordTemp";

    static {
        initialData();
    }

    /**
     * 初始化
     * 检测线程: 检测word进程健康状况
     * 清理线程: 清理转换产生的临时文件
     * 转换线程组: 执行word转换pdf
     */
    private static void initialData() {
        try {
            log.info("启用Com多线程支持");
            ComThread.InitMTA();
            log.info("初始化转换线程组");
            Runtime.getRuntime().exec("taskkill /F /IM WINWORD.EXE");
            Thread.sleep(200);
            log.info("初始化清理线程");
            executor.execute(JacobOffice::runCleaner);
            log.info("初始化检测线程");
            executor.execute(JacobOffice::runAlive);
            for (int i = 0; i < processorNum * 2; i++) {
                executor.execute(JacobOffice::runConvert);
            }


        } catch (Exception e) {
            log.error("初始化word进程异常", e);
        }
    }

    /**
     * 创建word进程并记录pid
     *
     * @return
     */
    private static synchronized ActiveXComVo createApp() {
        ActiveXComponent app = new ActiveXComponent("Word.Application");
        Set<String> pids = getWinwordPid("WINWORD.EXE");
        for (String pid : pids) {
            if (wordApps.stream().noneMatch(activeXComVo -> pid.equals(activeXComVo.getWordPid()))) {
                log.info("创建word进程,pid:{}", pid);
                ActiveXComVo activeXComVo = new ActiveXComVo();
                activeXComVo.setApp(app).setWordPid(pid).setLastUpdateTime(System.currentTimeMillis()).setCurrentThreadName(Thread.currentThread().getName());
                return activeXComVo;
            }
        }
        return null;
    }

    /**
     * 线程run方法
     * 从队列中获取word进行转换成pdf
     */
    private static void runConvert() {
        ActiveXComVo activeXComVo = createApp();
        wordApps.add(activeXComVo);
        while (!Thread.interrupted()) {
            ConvertedTarget target = targets.poll();
            try {
                if (!ObjectUtils.isEmpty(target)) {
                    ActiveXComponent app = activeXComVo.getApp();
                    //检查word是否正确加载签章插件v9
                    if (checkOffice360AddIn(app) < 0) {
                        log.info("检测到word进程没有加载签章插件");
                        targets.offer(target);
                        //休眠100s等待超时回收线程
                        Thread.sleep(100 * 1000);
                    }
                    Dispatch documents = app.getProperty("Documents").toDispatch();
                    app.setProperty("Visible", false);
                    //打开word文档
                    Dispatch document = Dispatch.call(documents, "Open", target.getWord().getPath(), true, false).toDispatch();
                    Dispatch.call(document, "Activate");
                    Dispatch.call(document, "SaveAs", target.getPdf().getPath(), new Variant(17));
                    Dispatch.call(document, "Close");

                    if (target.getPdf().exists()) {
                        //文件转换成功
                        //唤醒转换pdf的调用者
                        target.getCountDownLatch().countDown();
                        log.info("文件转换成功");
                    } else {
                        target.setTryCount(target.getTryCount() + 1);
                        targets.offer(target);
                    }
                } else {
                    Thread.sleep(200);
                }
                activeXComVo.setLastUpdateTime(System.currentTimeMillis());
            } catch (InterruptedException ie) {
                log.error("休眠线程收到中断信号");
                break;
            } catch (Exception e) {
                log.error("转换未知异常", e);
                log.info("文件转换失败");
                if (target.getTryCount() < 3) {
                    target.setTryCount(target.getTryCount() + 1);
                    //windows文件别之前的进程占用,需要重新拷贝一份
                    try {
                        String wordPath = target.getWord().getPath();
                        String sufix = wordPath.substring(wordPath.lastIndexOf("."));
                        String prefix = TMP_DIR + File.separatorChar + UUID.randomUUID().toString();
                        File word = new File(prefix + sufix);
                        byte[] bytes = new BASE64Decoder().decodeBuffer(target.getBase64File());
                        FileCopyUtils.copy(bytes, word);
                        target.setWord(word);
                    } catch (IOException ex) {
                        log.error("再次尝试转换创建文件失败", ex);
                    }
                    log.error("转换文件异常,重新进入队列,文件位置:{}", target.getWord().getPath());
                    targets.offer(target);
                }
            }
        }
    }


    private static void runAlive() {
        while (true) {
            try {
                log.info("进行word进程健康状况检测...");
                Long now = System.currentTimeMillis();
                for (ActiveXComVo activeXComVo : wordApps) {
                    Long alive = activeXComVo.getLastUpdateTime();
                    Long stay = now - alive;
                    log.info("当前word进程pid:{},活跃值:{}", activeXComVo.getWordPid(), stay);
                    if (stay > 20 * 1000) {
                        String pid = activeXComVo.getWordPid();
                        log.info("发现僵死转换线程,pid:{},活跃值:{} ms", pid, (now - alive));
                        String cmd = "taskkill /F /PID " + pid;
                        Runtime.getRuntime().exec(cmd);
                        log.info("结束winword.exe进程完毕");
                        killThreadByName(activeXComVo.getCurrentThreadName());
                        log.info("结束僵死线程完毕");
                        //队列中移除app
                        wordApps.remove(activeXComVo);
                        log.info("创建新线程替换僵死线程");
                        executor.execute(JacobOffice::runConvert);
                    }
                }
                Thread.sleep(5 * 1000);
            } catch (Exception e) {
                log.error("进行keepAlive时发现异常", e);
            }
        }
    }

    /**
     * 清理word临时文件
     */
    private static void runCleaner() {
        File tmpDir = new File(TMP_DIR);
        tmpDir.mkdirs();
        while (true) {
            try {
                Thread.sleep(6 * 60 * 60 * 1000);
                log.info("执行临时文件清理线程");
                File[] files = tmpDir.listFiles();
                for (File file : files) {
                    file.delete();
                }
            } catch (Exception e) {
                log.error("清理临时文件出现异常", e);
            }
        }
    }

    /**
     * 获取word进程pid列表
     *
     * @param processName
     * @return
     */
    private static Set<String> getWinwordPid(String processName) {
        Set<String> result = new HashSet<>();
        try {
            String cmd = "tasklist /V /FO CSV /FI \"IMAGENAME eq " + processName + "\"";
            log.info("获取系统运行的进程,{}", processName);
            BufferedReader br = new BufferedReader(new InputStreamReader(Runtime.getRuntime().exec(cmd).getInputStream(), Charset.forName("GBK")));
            String line;
            while ((line = br.readLine()) != null) {
                if (line != null && line.split(",")[0].contains(processName)) {
                    result.add(StringUtils.replace(line.split(",")[1], "\"", ""));
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return result;
    }

    /**
     * 检查word签章插件是否加载
     *
     * @param app
     * @return
     */
    private static int checkOffice360AddIn(ActiveXComponent app) {
        Dispatch office360Obj = null;
        try {
            Dispatch varComAddIns = Dispatch.get(app, "COMAddIns").toDispatch();
            int count = Dispatch.get(varComAddIns, "Count").getInt();
            for (int i = 1; i <= count; i++) {
                Dispatch addIn = Dispatch.call(varComAddIns, "Item", new Variant(i)).toDispatch();
                String ProgId = Dispatch.get(addIn, "ProgId").toString();
                if ("iSignatureOffice360.AddIn".equals(ProgId)) {
                    office360Obj = Dispatch.get(addIn, "Object").toDispatch();
                    log.info("word 签章插件加载正常");
                    break;
                }
            }
            if (office360Obj != null) {
                log.info("当前word进程签章插件正常加载");
                // 获取所有签章数量
                return 1;
            } else {
                log.error("获取签章客户端控件异常");
                return -1;
            }
        } catch (Exception e) {
            log.error("获取签章插件出现错误", e);
            return -1;
        }
    }

    private static boolean killThreadByName(String name) {
        Thread[] threads = findAllThread();
        for (Thread thread : threads) {
            if (thread.getName().equalsIgnoreCase(name)) {
                thread.interrupt();
                log.info("调用interrupt中断线程:" + name);
                return true;
            }
        }
        return false;
    }

    private static Thread[] findAllThread() {
        ThreadGroup currentGroup = Thread.currentThread().getThreadGroup();
        while (currentGroup.getParent() != null) {
            // 返回此线程组的父线程组
            currentGroup = currentGroup.getParent();
        }
        //此线程组中活动线程的估计数
        int noThreads = currentGroup.activeCount();
        Thread[] lstThreads = new Thread[noThreads];
        //把对此线程组中的所有活动子组的引用复制到指定数组中。
        currentGroup.enumerate(lstThreads);

//        for (Thread thread : lstThreads) {
//            System.out.println("线程数量：" + noThreads + " 线程id：" + thread.getId() + " 线程名称：" + thread.getName() + " 线程状态：" + thread.getState());
//        }
        return lstThreads;
    }

    public static void main(String[] args) throws Exception {
        System.out.println(Base64.getEncoder().encodeToString(FileCopyUtils.copyToByteArray(new File("F:\\金格签章,签章doc转pdf\\v9_2.doc"))));
    }

}
