
import graphviz
import excel_parser
import os
path = "Stammbaum_anonymous.xlsx"  # define path and filename for Excel file with  tree data


dot_node_size = str(0.08)  # size of the small dot that is plotted between two people



#Збільшення відстаней між вузлами та рівнями для кращого простору
graph_attributes = {
    'splines': "ortho",  # ортогональні лінії
    'nodesep': '2.0',  # Відстань між вузлами (збільшена)
    'ranksep': '2.0',  # Відстань між рівнями
    'overlap': 'false',  # Виключення перекриттів
    'newrank': 'true',  # Збереження нових рівнів
    'size': '25,20!',  # Дозволити графу розширювати ширину
    'margin': '0.5',
    'mindist': '2.0',  # Мінімальна відстань для краю
    'rankdir': 'TB',  # Орієнтація графа зверху вниз
}


graph_attributes2 = {
    "rank": "same",  # встановлення однакового рангу для групи вузлів
    "newrank": "true",
    "rankdir": "TB",
    "ranksep": "1.0",
    "nodesep": "0.8",  # збереження нових рівнів
}

node_attributes = {
    'style': 'filled',  # заповненість вузлів
    'shape': 'box',  # форма вузлів - прямокутник
    "label": "\\N",  # текстова позначка для вузлів
    "width": "2.5",  # ширина вузлів
    "height": "2.5",  # висота вузлів
    "fixedsize": "true",  # фіксовані розміри вузлів
    "penwidth": "4"  # товщина лінії
}

edge_attributes = {
    'dir': 'none',  # без стрілок на зв'язках
    'arrowhead': 'none',  # без стрілок на кінцях зв'язків
    "penwidth": "3"  # товщина лінії зв'язку
}






def parse_excel_data(excel_path):
    """
    Читає вміст з файлу Excel і розбирає його у зручний для використання формат.
    @param excel_path: Відносний шлях до файлу Excel.
    @return: Дані з файлу Excel у вигляді великого списку.
    """
    # read data from file
    family_tree_data = excel_parser.read_excel_content(path)
    # filter out "None" fields, replace with ""
    for i in range(0,len(family_tree_data)-1, 1):
        for j in range(0, len(family_tree_data[i]), 1):
            for k in range(0, len(family_tree_data[i][j]),1):
                if family_tree_data[i][j][k] == "None":
                    family_tree_data[i][j][k] = ""
    # add a field that indicates the hierarchy of that person
    family_tree_data[0].append([0 for i in range(0,len(family_tree_data[0][0]),1)])
    family_tree_data[0][-1][0] = "Hierarchy"
    # crop date fields to remove time and only leave date # 3, 5, 3
    family_tree_data[0][3] = [family_tree_data[0][3][i].split(" ")[0] for i in range(0,len(family_tree_data[0][3]),1)]
    family_tree_data[0][5] = [family_tree_data[0][5][i].split(" ")[0] for i in range(0,len(family_tree_data[0][3]),1)]
    family_tree_data[1][3] = [family_tree_data[1][3][i].split(" ")[0] for i in range(0,len(family_tree_data[1][3]),1)]
    return family_tree_data  # return list with data


def generate_flat_list(data):
    """
    Створює плаский список для всіх людей з даних Excel.
    @param data: Дані генеалогічного дерева, прочитані з файлу Excel.
    @return: плоский список із записами для кожної особи.
    """

    flat_list = []  # initialize return value
    for i in range(1, len(data[0][0]), 1):  # loop over all people
        flat_list.append([data[0][j][i] for j in range(0, len(data[0])-1, 1)])
        flat_list[-1][0] = int(flat_list[-1][0])  # convert own id to int
        flat_list[-1][7] = [int(flat_list[-1][7])] if "-" not in flat_list[-1][7] and len(flat_list[-1][7])>0 else []  # convert mother id to int
        flat_list[-1][8] = [int(flat_list[-1][8])] if "-" not in flat_list[-1][8] and len(flat_list[-1][8])>0 else []  # convert father id to int
        # convert child ids to int list
        flat_list[-1][10] = [int(child) for child in flat_list[-1][10].replace("-","").split(",") if len(child) > 0]
        flat_list[-1][9] = []  # initialize spouses as numeric list
    for i in range(0, len(flat_list), 1):  # populate spouse ids
        for j in range(1, len(data[1][0]), 1):  # loop over all spouses
            if int(flat_list[i][0]) == int(data[1][1][j]):  # add spouse 1
                flat_list[i][9].append(int(data[1][2][j]))
            elif int(flat_list[i][0]) == int(data[1][2][j]):  # add spouse 2
                flat_list[i][9].append(int(data[1][1][j]))
    return flat_list


def insert_into_flat_cluster(person, cluster):
    """
    Рекурсивно вставляє нову особу в заданий кластер.
    @param person: Локальний кластер поточної активної особи.
    @cluster: Поточний активний кластер, до якого буде додано особу.
    @return: Новий кластер з новою особою, прапорець для позначення успіху. 
    
    """
    new_cluster, flag = cluster.copy(), False  # initialize new cluster and flag
    # step 1: check people directly
    for i in range(1, 3, 1):  # check mother and father directly
        if new_cluster[i] and not flag:  # check if there is an entry and if the flag is not set
            if type(new_cluster[i][0]) != list:  # if this entry is not a cluster
                if person[0] == new_cluster[i][0]:  # if the ids are identical
                    new_cluster[i][0] = person  # assign person to cluster
                    flag = True  # set flag
    for h in range(4, 2, -1):  # check children and spouses directly
        if flag:  # skip rest if flag is set
            break
        for i, entry in enumerate(new_cluster[h]):  # check people directly
            if entry and not flag:  # check if there is an entry and if the flag is not set
                if type(entry) != list:  # if this entry is not a cluster
                    if person[0] == entry:  # if the ids are identical
                        new_cluster[h][i] = person  # assign person to cluster
                        flag = True  # set flag
            elif flag:  # exit loop if flag is set
                break
    # check parents, children and spouses recursively
    for i, category in enumerate(new_cluster[1:]):  # loop over all entries
        if flag:  # check exit condition
            break
        for j, entry in enumerate(category):  # loop over all indices in current category
            if flag:  # check exit condition
                break
            if type(entry) == list:  # check if current entry is not a cluster
                # recursive call
                new_cluster[i+1][j], flag = insert_into_flat_cluster(person, entry)
    return new_cluster, flag  # return the potentially updated cluster and the flag




def generate_flat_master_cluster(flat_list):
    """
    Перетворює плаский список усіх людей на плаский список даних кластерів.
    @param flat_list: Плоский список всіх людей з усіма даними з файлу Excel.
    @return: Плоский список, що містить релевантні кластеру дані для всіх людей
    """
    # initialize list to indicate if the person was found + return value
    person_found, cluster = [False for _ in range(0, len(flat_list), 1)], []
    for i, person in enumerate(flat_list):  # loop over all people
        # generate entry for current person: self, mother, father, spouses, children
        current = [person[0], person[7], person[8], person[9], person[10]]
        if not cluster:  # handle first cycle
            cluster = current  # assign current node as cluster root
            person_found[i] = True
        else:  # applies for all other cases
            cluster, person_found[i] = insert_into_flat_cluster(current, cluster)
    # sometimes the order of parents and children is swapped. Go over list again
    person_found_prev, flag = person_found.copy(), True
    while flag:  # loop until flag is reset
        for i, person in enumerate(flat_list):  # loop over all people
            if not person_found[i]:  # check if current person was already found
                # generate entry for current person: self, mother, father, spouses, children
                current = [person[0], person[7], person[8], person[9], person[10]]
                # try to fit current person into cluster
                cluster, person_found[i] = insert_into_flat_cluster(current, cluster)
        if person_found == person_found_prev or False not in person_found:
            flag = False  # exit the loop due to either no change or all people found
        person_found_prev = person_found.copy()  # copy list for next loop
    return cluster  # return value


# https://stackoverflow.com/questions/71571613/implement-family-tree-visualization-in-graphviz
def add_child_node(tree, parent1, parent2, child, person_plotted, flat_list):
    # Крок 1: Отримання всіх дітей
    children1 = [] if not parent1 else [child for child in parent1[10]]
    for i, child in enumerate(children1):
        if type(child) != int:
            children1[i] = child[0]
    children2 = [] if not parent2 else [child for child in parent2[10]]
    for i, child in enumerate(children2):
        if type(child) != int:
            children2[i] = child[0]
    
    children = [flat_list[child_id] for child_id in children1 if child_id in children2] if parent1 and parent2 else \
               [flat_list[c] for c in children1] if parent1 else [flat_list[c] for c in children2]

    # Сортування дітей за роком народження
    try:
        children = sorted(children, key=lambda x: int(x[3].split("-")[0]) if x[3] and "-" in x[3] else float('inf'))
    except ValueError as e:
        print(f"Помилка сортування дітей: {e}")
        children = []

    # Крок 2: Додавання вузлів батьків
    person_plotted, parent_node = add_spouse_node(tree, parent1, parent2, person_plotted)
    p1_id = -1 if not parent1 else parent1[0]
    p2_id = -1 if not parent2 else parent2[0]

    # Крок 3: Додавання єдиного вузла (центру) для дітей
    center_node = f"CENTER_{p1_id}_{p2_id}"
    tree.node(center_node, shape="point", width=str(dot_node_size))
    tree.edge(parent_node, center_node)

    # Крок 4: Додавання точок для дітей
    child_nodes = []
    with tree.subgraph(graph_attr=graph_attributes2) as sub_tree:
        for child in children:
            child_node = f"N{p1_id}_{p2_id}_{child[0]}"
            sub_tree.node(child_node, shape="point", width=str(dot_node_size))
            child_nodes.append(child_node)

    # З'єднання центру з кожною дитиною
    for child_node in child_nodes:
        tree.edge(center_node, child_node)

    # Крок 5: Додавання дітей до дерева
    with tree.subgraph(graph_attr=graph_attributes2) as sub_tree:
        for i, child in enumerate(children):
            child_node = str(child[0])
            if child[0] not in person_plotted or not person_plotted[child[0]]:
                label, args = generate_node_arguments(child)
                sub_tree.node(child_node, label, **args)
                person_plotted[child[0]] = True
            tree.edge(child_nodes[i], child_node)

    return person_plotted




def connect_to_center_or_middle(tree, parent_node, child_nodes):
    """
    З'єднує вузол батьків із середнім вузлом для їхніх дітей.
    @param tree: Дерево, яке обробляється.
    @param parent_node: Вузол батьків.
    @param child_nodes: Список вузлів дітей.
    """
    if len(child_nodes) % 2 == 0:
        mid_left = child_nodes[len(child_nodes) // 2 - 1]
        mid_right = child_nodes[len(child_nodes) // 2]
        center_node = f"CENTER_{parent_node}"
        tree.node(center_node, shape="point", **{"width": dot_node_size})
        tree.edge(parent_node, center_node)
        tree.edge(center_node, mid_left)
        tree.edge(center_node, mid_right)
    else:
        mid = child_nodes[len(child_nodes) // 2]
        tree.edge(parent_node, mid)




def generate_node_arguments(person):
    """
    Допоміжна функція для генерації аргументів вузла для вказаної особи.
    @param person: Особа, для якої будуть згенеровані аргументи.
    @return: Текстова мітка та аргументи для вузла person.
    """
    birth = "?" if len(person[3]) < 4 else person[3]  # generate birth string
    death = person[5]  # generate death string
    if len(death) < 4 and person[12]:
        death = "today" if person[12] else "?"
    
    # generate HTML-based labels for people nodes
    full_name = f"{person[1]} {person[2]}"
    if len(full_name) > 22:
        # figure out where to split the name
        split_index = 22
        while full_name[split_index] != " ":
            split_index -= 1
        name1 = full_name[0:split_index]
        name2 = full_name[split_index:]
        # handle long names with bold and color
        label = f"<{name1}<BR/>{name2}<BR/><FONT POINT-SIZE=\"9\" COLOR=\"#003366\"><B>"
        label += f"{birth} - {death}</B></FONT><BR/><FONT POINT-SIZE=\"5\"> </FONT><BR ALIGN=\"CENTER\"/>>"
    else:
        label = f"<{person[1]} {person[2]}<BR/><BR/><FONT POINT-SIZE=\"9\" COLOR=\"#003366\"><B>{birth}"
        label += f" - {death}</B></FONT><BR/><FONT POINT-SIZE=\"5\"> </FONT><BR ALIGN=\"CENTER\"/>>"
    
    # add constant node arguments
    node_arguments = {
        "height": str(2.6),
        "width": str(2),
        "penwidth": str(3),
        "fixedsize": "true",
        "imagepos": "tc",
        "imagescale": "true",
        "labelloc": "bc",
        "group": f"G{person[0]}"
    }

    # determine the image path
    default_folder = "Images"
    image_path = os.path.join(default_folder, f"{person[0]}.png")
    if not os.path.exists(image_path):
        # If ID photo doesn't exist, assign default based on gender
        image_path = os.path.join(default_folder, "man.png" if person[11] == "m" else "woman.png")
    node_arguments["image"] = image_path

    # add node color based on gender and transparency
    node_arguments["fillcolor"] = "#D3D3D3"  # Light gray background for text
    node_arguments["style"] = "filled"  # Ensure that the background color is applied

    # Make sure image background is transparent
    node_arguments["shape"] = "rect"
    node_arguments["imagealign"] = "center"

    return label, node_arguments  # return the node label and additional arguments



# https://stackoverflow.com/questions/71571613/implement-family-tree-visualization-in-graphviz
def add_spouse_node(tree, person1, person2, person_plotted):
    """
    Додає подружжя та зв'язок між подружжям до поточного дерева.
    @param дерево: Поточне активне дерево.
    @param person1: Особа 1 у зв'язку.
    @param person2: Особа 2 у зв'язку.
    @param person_plotted: Допоміжні прапорці, які вказують, чи існує особа на дереві.
    @return: Оновлена версія person_plotted, назва додаткового вузла.
    """
    # TODO: HANDLE SINGLE PARENT
    attr = graph_attributes2.copy()
    attr["peripheries"] = "0"
    p1_id = -1 if not person1 else person1[0]
    p2_id = -1 if not person2 else person2[0]
    # generate a sub graph for spouse pair
    with tree.subgraph(name=f"cluster_{p1_id}_{p2_id}", graph_attr = attr) as sub_tree:
        # generate labels and node arguments
        if person1 and not person_plotted[person1[0]]:  # check if person 1 exists on graph
                label1, arguments1 = generate_node_arguments(person1)
                sub_tree.node(str(person1[0]), label1, **arguments1)
                person_plotted[person1[0]] = True
        if person2 and not person_plotted[person2[0]]:  # check if person 2 exists on graph
            label2, arguments2 = generate_node_arguments(person2)
            sub_tree.node(str(person2[0]), label2, **arguments2)
            person_plotted[person2[0]] = True
        # add node that connects both people. Important for children
        node_name = f"N{p1_id}_{p2_id}"
        sub_tree.node(node_name, shape="point", **{"width":str(0.08)})
        # generate edges between people nodes and connector node
        if person1:
            sub_tree.edge(str(person1[0]), node_name)
        if person2:
            sub_tree.edge(node_name, str(person2[0]))
    return person_plotted, node_name  # return list of plotted people and extra node name


def plot_next_person(flat_list, tree, cluster_flat, person_plotted, person_parsed):
    """
    Рекурсивна функція для додавання наступної людини з плаского списку до заданого кластера.
    @param flat_list: Плоский список, що містить всіх людей, які будуть додані до кластера.
    @param tree: Активне на даний момент дерево.
    @param cluster_flat: Плоский список, що містить всі дані, які стосуються кластера.
    @param person_plotted: Допоміжний прапорець, який вказує, чи було додано особу до дерева.
    @param person_parsed: Допоміжний прапорець, який вказує, чи була особа рекурсивно викликана через цю функцію.
    @return: Оновлена версія функції person_plotted.

    """
    if type(cluster_flat) == int:  # handle szenario where cluster is only one index
        if not person_parsed[cluster_flat]:  # set flag to true if not alredy
            person_parsed[cluster_flat] = True
        return person_plotted  # return before rest of code
    if cluster_flat[3]:  # check if current person has any spouses
        for spouse in cluster_flat[3]:  # add all spouses as new nodes
            if type(spouse) != list:
                spouse = [spouse]
            if person_plotted[spouse[0]]:  # skip already plotted people
                continue
            # generate person information lists
            person1 = flat_list[cluster_flat[0]]
            person2 = flat_list[spouse[0]]
            # add the spouse node
            person_plotted, _ = add_spouse_node(tree, person1, person2, person_plotted)
    if cluster_flat[4]:  # check if current person has any children
        for child in cluster_flat[4]:  # add all children as new nodes
            if type(child) != list:
                child = [child, [-1], [-1]]
            if person_plotted[child[0]]:  # skip already plotted people
                continue
            # generate person information lists
            try:
                parent1 = flat_list[child[1][0][0]]
            except:
                if child[1] == []:
                    child[1] = [-1]
                parent1 = flat_list[child[1][0]]
            try:
                parent2 = flat_list[child[2][0][0]]
            except:
                if child[2] == []:
                    child[2] = [-1]
                parent2 = flat_list[child[2][0]]
            # add the child node
            person_plotted = add_child_node(tree, parent1, parent2, child, person_plotted, flat_list)
    if cluster_flat[1] or cluster_flat[2]:  # check if current person has parents
        parent1, parent2, parent_index1, parent_index2 = [], [], -1, -1
        if type(cluster_flat[1]) != list:
            cluster_flat[1] = [[cluster_flat[1]]]
        if type(cluster_flat[2]) != list:
            cluster_flat[2] = [[cluster_flat[2]]]
        if cluster_flat[1]:
            parent_index1 = cluster_flat[1][0] if type(cluster_flat[1][0]) == int else cluster_flat[1][0][0]
        if cluster_flat[2]:
            parent_index2 = cluster_flat[2][0] if type(cluster_flat[2][0]) == int else cluster_flat[2][0][0]
        if parent_index1 >= 0 and not person_plotted[parent_index1]:
            parent1 = flat_list[parent_index1]
        if parent_index2 >= 0 and not person_plotted[parent_index2]:
            parent2 = flat_list[parent_index2]
        if parent1 or parent2:  # add the parent nodes if valid
            person_plotted = add_child_node(tree, parent1, parent2, 
                                               flat_list[cluster_flat[0]],
                                               person_plotted, flat_list)
    # set the flag for this person
    person_parsed[cluster_flat[0]] = True
    #"""
    # plot next people
    if cluster_flat[3]:  # spouses
        for spouse in cluster_flat[3]:
            if type(spouse) == list and not person_parsed[spouse[0]]:
                plot_next_person(flat_list, tree, spouse, person_plotted, person_parsed)
    #"""
    if cluster_flat[4]:  # children
        for child in cluster_flat[4]:
            if type(child) == list and not person_parsed[child[0]]:
                plot_next_person(flat_list, tree, child, person_plotted, person_parsed)
    #"""
    if cluster_flat[1]:  # mother
        plot_next_person(flat_list, tree, cluster_flat[1][0], person_plotted, person_parsed)
    if cluster_flat[2]:  # father
        plot_next_person(flat_list, tree, cluster_flat[2][0], person_plotted, person_parsed)
    #"""
    return person_plotted


def plot_flat_master_cluster(flat_list, master_cluster_flat):
    """
    Будує графік заданого плоского головного кластера за допомогою graphviz. Додаткова інформація
    береться з першого аргументу, який є пласким списком усіх людей.
    Кластер містить лише індекси, плаский список пов'язує індекси з 
    особистими даними.
    @param flat_list: Плоский список, що містить всі дані для кожної особи.
    @param master_cluster_flat: Плоский кластер, що містить всі дані, які стосуються кластера.
    @return: Деревовидний об'єкт, що містить всіх людей з плаского списку.
    """
    person_plotted = [False for _ in range(0, len(flat_list), 1)]  # helper variable
    person_parsed = [False for _ in range(0, len(flat_list), 1)]  # helper variable
    # generate basic tree with general information
    tree = graphviz.Graph(engine='dot',
                          graph_attr=graph_attributes,
                          node_attr=node_attributes,
                          edge_attr=edge_attributes,
                          encoding='utf8',
                          filename='family_tree',
                          format='svg')
    # --- handle the first node ---
    person_plotted = plot_next_person(flat_list, tree, master_cluster_flat, person_plotted, person_parsed)
    # save tree to disk
    tree.save("family_tree")
    tree.view()  # show the tree
    return tree

# --- PROGRAM START ---

# step 1: get data from Excel file
family_tree_data = parse_excel_data(path)
# step 2: generate a flat list to look up data later
flat_list = generate_flat_list(family_tree_data)
# step 3: generate the master cluster with only indices
master_cluster_flat = generate_flat_master_cluster(flat_list)
# step 4: plot master cluster
tree = plot_flat_master_cluster(flat_list, master_cluster_flat)
