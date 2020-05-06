using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop;


namespace RandomList
{
    class Program
    {
        static void Main(string[] args)
        {
            while (true)
            {
                try
                {
                    string currentDirectory = Directory.GetCurrentDirectory();

                    string studentFileName = "estudiantes.txt";
                    Console.Write("Nombre del archivo de Estudiantes: ");
                    studentFileName = Console.ReadLine();
                    string studentPath = currentDirectory + "\\" + studentFileName;
                    if (File.Exists(studentPath))
                    {
                        while (true)
                        {
                            string themeFileName = "temas.txt";
                            Console.Write("Nombe del archivo de Temas: ");
                            themeFileName = Console.ReadLine();
                            string themePath = currentDirectory + "\\" + themeFileName;

                            if (File.Exists(themePath))
                            {

                                Console.Write("Inserte la cantidad de estudiantes por grupos: ");
                                int cantEstGroup = int.Parse(Console.ReadLine());

                                List<Grupo> groups = new List<Grupo>();
                                List<Estudiante> students = new List<Estudiante>();
                                List<Tema> themes = new List<Tema>();

                                using (StreamReader reader = new StreamReader(studentPath))
                                {
                                    string line;
                                    Estudiante student;
                                    while ((line = reader.ReadLine()) != null)
                                    {
                                        student = new Estudiante();
                                        student.FullName = line;
                                        students.Add(student);
                                    }
                                }

                                using (StreamReader reader = new StreamReader(themePath))
                                {
                                    string line;
                                    Tema theme;
                                    while ((line = reader.ReadLine()) != null)
                                    {
                                        theme = new Tema();
                                        theme.Name = line;
                                        themes.Add(theme);
                                    }
                                }

                                if (students.Count >= cantEstGroup && themes.Count >= cantEstGroup)
                                {
                                    int groupsCount = students.Count / cantEstGroup;
                                    int rest = students.Count % cantEstGroup;
                                    int themesCount = themes.Count / groupsCount;
                                    int restThemes = themes.Count % groupsCount;
                                    if (students.Count >= groupsCount && themes.Count >= groupsCount)
                                    {
                                        Random random = new Random();
                                        Grupo group;
                                        List<Estudiante> studentList;
                                        List<Tema> themeList;
                                        for (int a = 0; a < groupsCount; a++)
                                        {
                                            group = new Grupo();
                                            group.Nro = a + 1;
                                            studentList = new List<Estudiante>();
                                            themeList = new List<Tema>();
                                            for (int b = 0; b < cantEstGroup; b++)
                                            {
                                                int i = random.Next(0, students.Count);
                                                studentList.Add(students[i]);
                                                students.RemoveAt(i);
                                            }
                                            for (int c = 0; c < themesCount; c++)
                                            {
                                                int i = random.Next(0, themes.Count);
                                                themeList.Add(themes[i]);
                                                themes.RemoveAt(i);
                                            }
                                            group.Estudiantes = studentList;
                                            group.Temas = themeList;
                                            groups.Add(group);
                                        }
                                        List<int> counts = new List<int>();
                                        bool valid = false;
                                        while (rest > 0)
                                        {
                                            valid = false;
                                            int i = random.Next(0, groups.Count);
                                            if (counts.Count > 0 && groups.Count > 1)
                                            {
                                                while (!valid)
                                                {
                                                    i = random.Next(0, groups.Count);
                                                    foreach (var item in counts)
                                                    {
                                                        valid = true;
                                                        if (i == item)
                                                        {
                                                            valid = false;
                                                            break;
                                                        }
                                                    }
                                                }
                                            }
                                            counts.Add(i);
                                            if (counts.Count == groups.Count)
                                                counts.Clear();
                                            int x = random.Next(0, students.Count);
                                            List<Estudiante> studentsTemp = groups[i].Estudiantes;
                                            Estudiante studentTemp = students[x];

                                            studentsTemp.Add(studentTemp);
                                            groups[i].Estudiantes = studentsTemp;
                                            students.RemoveAt(x);
                                            rest--;
                                        }
                                        counts.Clear();
                                        while (restThemes > 0)
                                        {
                                            int i = random.Next(0, groups.Count);
                                            valid = false;
                                            if (counts.Count > 0 && groups.Count > 1)
                                            {
                                                while (!valid)
                                                {
                                                    i = random.Next(0, groups.Count);
                                                    foreach (var item in counts)
                                                    {
                                                        valid = true;
                                                        if (i == item)
                                                        {
                                                            valid = false;
                                                            break;
                                                        }
                                                    }
                                                }
                                            }
                                            counts.Add(i);
                                            if (counts.Count == groups.Count)
                                                counts.Clear();
                                            int x = random.Next(0, themes.Count);
                                            List<Tema> themesTemp = groups[i].Temas;
                                            Tema themeTemp = themes[x];

                                            themesTemp.Add(themeTemp);
                                            groups[i].Temas = themesTemp;
                                            themes.RemoveAt(x);
                                            restThemes--;
                                        }
                                        string fileName = "Resultado-1.txt";
                                        string filePath = currentDirectory + "\\" + fileName;
                                        if (File.Exists(filePath))
                                        {
                                            string[] divition = fileName.Split('-');
                                            divition[1] = divition[1].Split('.')[0];
                                            int numberFile = int.Parse(divition[1]);
                                            numberFile += 1;
                                            filePath = currentDirectory + $"\\Resultado-{numberFile}.txt";
                                        }
                                        using (StreamWriter sw = File.AppendText(filePath))
                                        {
                                            sw.WriteLine("==========================================================================");
                                            foreach (var grp in groups)
                                            {
                                                sw.WriteLine($"Grupo #{grp.Nro}");
                                                sw.WriteLine("----------------------------------------------------------------------");
                                                int i = 1;
                                                foreach (var student in grp.Estudiantes)
                                                {
                                                    sw.WriteLine($"{i}. {student.FullName}");
                                                    i++;
                                                }
                                                sw.WriteLine("----------------------------------------------------------------------");
                                                i = 1;
                                                foreach (var theme in grp.Temas)
                                                {
                                                    sw.WriteLine($"Tema #{i}: {theme.Name}.");
                                                    i++;
                                                }
                                                sw.WriteLine("\n======================================================================");
                                            }
                                        }

                                        using (StreamReader sr = new StreamReader(filePath))
                                        {
                                            string line;
                                            while ((line = sr.ReadLine()) != null)
                                            {
                                                Console.WriteLine(line);
                                            }
                                        }
                                    }
                                    else
                                        Console.WriteLine("Hay mas grupos que estudiantes.");
                                }
                                else
                                    Console.WriteLine("Hay menos estudiantes o temas que el minimo requerido.");
                                break;
                            }
                            else
                                Console.WriteLine("Archivo no existe.");
                        }
                    }
                    else
                        Console.WriteLine("Archivo no existe");
                }
                catch(Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                Console.ReadKey();
            }
        }

    }
}
