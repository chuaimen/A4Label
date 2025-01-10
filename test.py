from PyQt5.QtWidgets import QApplication, QTreeView, QStandardItemModel, QStandardItem
from PyQt5.QtCore import QModelIndex

app = QApplication([])

# 创建模型
model = QStandardItemModel()
rootItem = QStandardItem("根节点")
model.appendRow(rootItem)
child1 = QStandardItem("子节点1")
child2 = QStandardItem("子节点2")
rootItem.appendRow(child1)
rootItem.appendRow(child2)

# 创建视图
tree = QTreeView()
tree.setModel(model)
tree.expandAll()


# 连接到选择变化信号
def on_selection_changed(selected, deselected):
    if selected.indexes():
        index = selected.indexes()[0]  # 假设只选中了一个项
        item = model.itemFromIndex(index)
        print(f"选中项的文本: {item.text()}")


tree.selectionModel().selectionChanged.connect(on_selection_changed)

tree.show()

app.exec_()