import javax.swing.*;

public class Main {
    public static void main(String[] args) {
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (IllegalAccessException | InstantiationException | UnsupportedLookAndFeelException
                | ClassNotFoundException e) {
            e.printStackTrace();
        }

        // GESTION FENÊTRE

        JFrame fenetre = new JFrame();
        fenetre.setResizable(false);

        Reader reader = new Reader(fenetre);
        // Définit un titre pour notre fenêtre
        fenetre.setTitle("catalogGenJAVA");
        // Termine le processus lorsqu'on clique sur la croix rouge
        fenetre.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        // Et enfin, la rendre visible
        fenetre.setVisible(true);
        // Instanciation d'un objet JPanel

        // Définition de sa couleur de fond
        // On prévient notre JFrame que notre JPanel sera son content pane
        fenetre.setContentPane(reader);
        fenetre.pack();
        fenetre.setVisible(true);
    }
}