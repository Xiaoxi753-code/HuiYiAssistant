<script setup>
defineProps({
    /** 待办角标数量 */
    count: { type: Number, default: 0 },
})

const emit = defineEmits(['open'])

function onClick() {
    emit('open')
}
</script>

<template>
    <button type="button" class="todo-bell" title="待办事项" aria-label="待办事项" @click="onClick">
        <span class="todo-bell__emoji" aria-hidden="true">🔔</span>
        <span v-if="count > 0" class="todo-bell__badge">{{ count > 99 ? '99+' : count }}</span>
    </button>
</template>

<style scoped>
.todo-bell {
    position: relative;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 38px;
    height: 38px;
    flex-shrink: 0;
    padding: 0;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    background: #fff;
    cursor: pointer;
    transition:
        background 0.15s ease,
        border-color 0.15s ease;
}

.todo-bell:hover {
    background: #f8fafc;
    border-color: #cbd5e1;
}

.todo-bell__emoji {
    font-size: 18px;
    line-height: 1;
    display: inline-flex;
    align-items: center;
    justify-content: center;
}

.todo-bell__badge {
    position: absolute;
    top: 2px;
    right: 2px;
    min-width: 16px;
    height: 16px;
    padding: 0 4px;
    border-radius: 999px;
    background: #ef4444;
    color: #fff;
    font-size: 0.65rem;
    font-weight: 700;
    line-height: 16px;
    text-align: center;
}
</style>
