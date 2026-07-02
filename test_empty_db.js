const { db } = require('./backend/db');

// Insert a fake empty subgroup
db.prepare("INSERT INTO subgroups (group_id, name, display_order) VALUES (1, 'Empty Subgroup', 999)").run();
// Insert a fake empty group
db.prepare("INSERT INTO groups (engagement_id, name, display_order) VALUES (1, 'Empty Group', 999)").run();

const engagement_id = 1;
const treeData = db.prepare(`
  SELECT 
    g.id AS group_id, g.name AS group_name, g.display_order AS group_order,
    sg.id AS subgroup_id, sg.name AS subgroup_name, sg.display_order AS subgroup_order,
    ssg.id AS sub_subgroup_id, ssg.name AS sub_subgroup_name, ssg.display_order AS sub_subgroup_order,
    (SELECT COUNT(*) FROM trial_balance_ledgers WHERE sub_subgroup_id = ssg.id) AS ssg_ledger_count
  FROM groups g
  LEFT JOIN subgroups sg ON sg.group_id = g.id
  LEFT JOIN sub_subgroups ssg ON ssg.subgroup_id = sg.id
  WHERE g.engagement_id = ?
  ORDER BY g.display_order ASC, sg.display_order ASC, ssg.display_order ASC
`).all(engagement_id);

const tree = [];
let currentGroup = null;
let currentSubgroup = null;

for (const row of treeData) {
  if (!currentGroup || currentGroup.id !== row.group_id) {
    currentGroup = {
      id: row.group_id,
      name: row.group_name,
      display_order: row.group_order,
      ledger_count: 0,
      subgroups: []
    };
    tree.push(currentGroup);
    currentSubgroup = null;
  }

  if (row.subgroup_id) {
    if (!currentSubgroup || currentSubgroup.id !== row.subgroup_id) {
      currentSubgroup = {
        id: row.subgroup_id,
        name: row.subgroup_name,
        display_order: row.subgroup_order,
        ledger_count: 0,
        sub_subgroups: []
      };
      currentGroup.subgroups.push(currentSubgroup);
    }

    if (row.sub_subgroup_id) {
      const ssg = {
        id: row.sub_subgroup_id,
        name: row.sub_subgroup_name,
        display_order: row.sub_subgroup_order,
        ledger_count: row.ssg_ledger_count
      };
      
      currentSubgroup.sub_subgroups.push(ssg);
      currentSubgroup.ledger_count += ssg.ledger_count;
      currentGroup.ledger_count += ssg.ledger_count;
    }
  }
}

// Find the ones we just added to see if they formatted properly
console.log(JSON.stringify(tree.find(g => g.name === "Empty Group"), null, 2));
console.log(JSON.stringify(tree.find(g => g.name === "Income").subgroups.find(sg => sg.name === "Empty Subgroup"), null, 2));

// Cleanup
db.prepare("DELETE FROM groups WHERE name = 'Empty Group'").run();
db.prepare("DELETE FROM subgroups WHERE name = 'Empty Subgroup'").run();

