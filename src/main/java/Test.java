public class Test {
    public static void main(String[] args) {
        String name = "85-7821-1 Пензель для фарбування";
        if (name.contains("Пензель")){
            System.out.println("wow");
        }



        if (name.contains("Фарба ") || name.contains("Краска ") || name.contains("Лак ") || name.contains("грунт ") || name.contains("Морілка ")) {
            System.out.println("Фарба");
//                    Інші
        } else if (name.contains("Балон") || name.contains("Диск") || name.contains("Стрічка") || name.contains("Пензель ") || name.contains("Пензлі") || name.contains("Частини") || name.contains("Пензлі")) {
            System.out.println("Балон");
//                    Хімія
        } else if (name.contains("Тексапон") || name.contains("Деріфат") || name.contains("Дехікварт") || name.contains("Трезаліт") || name.contains("Розчинник") || name.contains("Ларопал") || name.contains("Глюкопон") || name.contains("Трезоліт") || name.contains("Шпаклівка") || name.contains("ацетат") || name.contains("Дехітон") || name.contains("Тінувін") || name.contains("Трилон") || name.contains("Лютенсол") || name.contains("Отверджувач") || name.contains("Антигравій")) {
            System.out.println("Хімія");
        }


    }
}

