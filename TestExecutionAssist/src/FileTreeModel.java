package src;

import javax.swing.JFrame;
import javax.swing.JScrollPane;
import javax.swing.JTree;
import javax.swing.tree.TreeModel;
import javax.swing.tree.TreePath;
import javax.swing.event.TreeModelListener;
import javax.swing.event.TreeModelEvent;
import javax.swing.event.EventListenerList;

import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.util.List;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Map;
import java.util.HashMap;
import java.io.File;
import java.io.Serializable;
import java.io.ObjectOutputStream;
import java.io.ObjectInputStream;


/** Example of a simple static TreeModel. It contains a
    (java.io.File) directory structure.
    (C) 2001 Christian Kaufhold (ch-kaufhold@gmx.de)
*/

public class FileTreeModel
    implements TreeModel, Serializable, Cloneable
{
    protected EventListenerList listeners;

    private static final Object LEAF = new Serializable() { };
    
    private static List<File>  moduelesList = new ArrayList();
 
    private Map map;


    private File root;


    public FileTreeModel(File root)
    {
        this.root = root;

        if (!root.isDirectory())
            map.put(root, LEAF);

        this.listeners = new EventListenerList();

        this.map = new HashMap();
    }


    public Object getRoot()
    {
        return root;
    }

    public boolean isLeaf(Object node)
    {
        return map.get(node) == LEAF;
    }


    public int getChildCount(Object node)
    {
    	File fileSysEntity;
    	if(!node.toString().equals(root.getAbsolutePath())) {
    		//fileSysEntity = new File(root.getAbsoluteFile() + "\\" + node.toString());
    		fileSysEntity = (File) node;
    	}
    	else{
    		fileSysEntity = (File) node;
    	} 	
    	getmoduleList(fileSysEntity);
    	return moduelesList.size();
    }

    public Object getChild(Object parent, int index)
    {
    	File directory;
    	if(!parent.toString().equals(root.getAbsolutePath())) {
    		//directory = new File(root.getAbsoluteFile() + "\\" + parent.toString());
    		directory = (File) parent;
    	}
    	else{
    		directory = (File) parent;
    	}
    	 Object value = map.get(directory);
         if (value == LEAF)
             return null;
		 map.put(directory, moduelesList);
		 return moduelesList.get(index);
    }

    public int getIndexOfChild(Object parent, Object child)
    {
    	File directory;
    	if(parent.toString()!=root.getAbsolutePath()) {
    		//directory = new File(root.getAbsoluteFile() + "\\" + parent.toString());
    		directory = (File)parent;
    	}
    	else{
    		directory = (File) parent;
    	}
        File fileSysEntity = (File)child;
        String[] children = directory.list();
        int result = -1;

        for ( int i = 0; i < children.length; ++i ) {
            if ( fileSysEntity.getName().equals( children[i] ) ) {
                result = i;
                break;
            }
        }

        return result;
    }
    
    public void getmoduleList(File directory) {
    	moduelesList.clear();
    	for(File eachFile : directory.listFiles()) {
 	        if (eachFile.isDirectory()) {
 	        	if(new File(eachFile.getAbsolutePath() + "\\moduleDetails.ts").exists())  {
 	        		moduelesList.add(eachFile);
 	        	}
 	        	else if(new File(eachFile.getAbsolutePath() + "\\testDetails.tc").exists()) {
 	        		moduelesList.add(eachFile);
 	        		map.put(eachFile, LEAF);
 	        	}
 	        }
     	}
    }

    

     public void valueForPathChanged(final TreePath pPath, final Object pNewValue) {
        
    }

    public void addTreeModelListener(TreeModelListener l)
    {
        listeners.add(TreeModelListener.class, l);
    }

    public void removeTreeModelListener(TreeModelListener l)
    {
        listeners.remove(TreeModelListener.class, l);
    }

    public Object clone()
    {
        try
        {
            FileTreeModel clone = (FileTreeModel)super.clone();
    
            clone.listeners = new EventListenerList();
    
            clone.map = new HashMap(map);
    
            return clone;
        }
        catch (CloneNotSupportedException e)
        {
            throw new InternalError();
        }
    }
	
}