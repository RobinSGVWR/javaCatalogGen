public class ThreadBar implements Runnable {
    private Reader reader;
    private int value;

    public ThreadBar(Reader reader) {
        this.reader = reader;
    }

    public void run(int Value) {
        reader.getBar().setValue(value);
        reader.repaint();
        reader.revalidate();

        try {
            Thread.sleep(value);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
    }

    @Override
    public void run() {}
}