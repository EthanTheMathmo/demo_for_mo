"""
Implementation of doubly linked list for use in the undo function
"""


class doubllistNode:
    """
    self.parent = None if it has no parent, and self.child = None if it has no child
    """
    def __init__(self, parent, child, value):
        self.parent = parent
        self.child = child
        self.value = value    
    
    def del_descendants(self):
        if self.child == None:
            return
        else:
            self.child.del_descendants()
            del self.child
            self.child=None
            return
    
    def del_ancestors(self):
        if self.parent == None:
            pass
        else:
            self.parent.del_ancestors()
            del self.parent
            self.parent=None
            return
        
    def __repr__(self):
        repr_string = f"""doubllistNode instance:
        value = {self.value}
        Parent exists? {self.parent!=None}
        Child exists? {self.child!=None}
        """
        return repr_string
            
        

class doubllist:
    def __init__(self, current_node):
        self.current_node = current_node
        
    def step_forward(self):
        #moves the current node forward by one and returns True, returns False if the current node has no child node
        if self.current_node.child != None:
            self.current_node = self.current_node.child
            return True
        else:
            return False
        
    def step_backward(self):
        #moves the current node back by one and returns True, returns False if the current node has no parent node
        if self.current_node.parent != None:
            self.current_node = self.current_node.parent
            return True
        else:
            return False
        
    def insert_ahead(self, value=None):
        #adds a node ahead in the chain
        #if there is already a node ahead, we insert the node between them
        current_child = self.current_node.child
        if current_child == None:
            self.current_node.child = doubllistNode(parent=self.current_node, child=None, value=value)
        else:
            new_child = doubllistNode(parent=self.current_node, child=current_child, value=value)
            self.current_node.child = new_child
            
    def insert_behind(self, value=None):
        #adds a node behind in the chain
        #if there is already a node behind, we inset the node between them
        current_parent = self.current_node.parent
        if current_parent == None:
            self.current_node.parent = doubllistNode(parent=None, child=self.current_node, value=value)
        else:
            new_parent= doubllistNode(parent=current_parent, child=self.current_node, value=value)
            self.current_node.parent = new_parent    

    def __repr__(self):
        return_str = ""
        root_node = self.current_node
        while root_node.parent != None:
            root_node = root_node.parent
        node = root_node

        return_str += str(node)

        while node.child != None:
            node = node.child

            return_str += "-"*5+"\n" + str(node) 
        return return_str