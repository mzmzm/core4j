package minor.zzz.util.threadpool;

import minor.zzz.util.threadpool.support.CountLatch;

import java.util.concurrent.BlockingQueue;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.concurrent.TimeUnit;

/**
 * 适用于无法明确预估任务数同时又需要等待所有任务执行完毕的情况(比如任务是递归创建的)
 */
public class CountingThreadPoolExecutor extends ThreadPoolExecutor {

    protected final CountLatch numRunningTasks = new CountLatch(0);

    public CountingThreadPoolExecutor(int corePoolSize, int maximumPoolSize,
                                      long keepAliveTime, TimeUnit unit, BlockingQueue<Runnable> workQueue) {
        super(corePoolSize, maximumPoolSize, keepAliveTime, unit, workQueue);
    }

    @Override
    public void execute(Runnable command) {
        numRunningTasks.increment();
        super.execute(command);
    }

    @Override
    protected void afterExecute(Runnable r, Throwable t) {
        numRunningTasks.decrement();
        super.afterExecute(r, t);
    }

    /**
     * Awaits the completion of all spawned tasks.
     */
    public void awaitCompletion() throws InterruptedException {
        numRunningTasks.awaitZero();
    }

    /**
     * Awaits the completion of all spawned tasks.
     */
    public void awaitCompletion(long timeout, TimeUnit unit)
            throws InterruptedException {
        numRunningTasks.awaitZero(timeout, unit);
    }

    /**
     * 等待所有任务完成后关闭线程池
     */
    public void awaitShutdown() {
        boolean isRunning = true;

        try {
            // 等待直到任务全部完成
            while (isRunning && !Thread.currentThread().isInterrupted()) {
                try {
                    awaitCompletion();
                    isRunning = false;
                } catch (InterruptedException ignore) {
                    Thread.interrupted();			// 重置中断状态
                }
            }
        } finally {
            shutdown();
        }
    }
}

